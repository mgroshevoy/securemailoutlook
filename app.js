'use strict';

var item, $select, subject = '', subj, emails, peopleArray = [];

(function () {
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        item = Office.context.mailbox.item;
        item.saveAsync(function (result) {
            // Get token from client
            console.log(result);
            Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function (result) {
                if (result.status === "succeeded") {
                    var accessToken = result.value;
                    console.log(accessToken);
                    var restHost = Office.context.mailbox.restUrl;
                    console.log(restHost);
                    getContactsItem(accessToken);
                } else {
                    // Handle the error
                    console.error(result);
                    getTokenFromClient();
                }
            });
        });

        $(document).ready(function () {

            var REGEX_EMAIL = '([a-z0-9!#$%&\'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&\'*+/=?^_`{|}~-]+)*@' +
                '(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)';

            // Define select field options
            $select = $('#select-to').selectize({
                //persist: false,
                plugins: ['remove_button'],
                preload: true,
                maxItems: null,
                valueField: 'email',
                labelField: 'name',
                searchField: ['name', 'email'],
                options: peopleArray,
                render: {
                    item: function (item, escape) {
                        return '<div>' +
                            (item.name ? '<span class="name">' + escape(item.name) + '</span>' : '') +
                            (item.email ? '<span class="email">' + escape(item.email) + '</span>' : '') +
                            '</div>';
                    },
                    option: function (item, escape) {
                        var label = item.name || item.email;
                        var caption = item.name ? item.email : null;
                        return '<div>' +
                            '<span class="label">' + escape(label) + '</span>' +
                            (caption ? '<span class="caption">' + escape(caption) + '</span>' : '') +
                            '</div>';
                    }
                },
                createFilter: function (input) {
                    var match, regex;

                    // email@address.com
                    regex = new RegExp('^' + REGEX_EMAIL + '$', 'i');
                    match = input.match(regex);
                    if (match) return !this.options.hasOwnProperty(match[0]);

                    // name <email@address.com>
                    regex = new RegExp('^([^<]*)\<' + REGEX_EMAIL + '\>$', 'i');
                    match = input.match(regex);
                    if (match) return !this.options.hasOwnProperty(match[2]);

                    return false;
                },
                create: function (input) {
                    if ((new RegExp('^' + REGEX_EMAIL + '$', 'i')).test(input)) {
                        return {email: input};
                    }
                    var match = input.match(new RegExp('^([^<]*)\<' + REGEX_EMAIL + '\>$', 'i'));
                    if (match) {
                        return {
                            email: match[2],
                            name: $.trim(match[1])
                        };
                    }
                    console.error('Invalid email address.');
                    return false;
                }
            });

            //var buttonComponents = [];
            var ButtonElements = document.querySelectorAll(".ms-Button");
            for(var i = 0; i < ButtonElements.length; i++) {
                new fabric['Button'](ButtonElements[i], function(event) {
                });
            }

            // Define checkbox component
            var checkBoxComponents = [];
            var CheckBoxElements = document.querySelectorAll(".ms-CheckBox");
            for (var i = 0; i < CheckBoxElements.length; i++) {
                checkBoxComponents.push(new fabric['CheckBox'](CheckBoxElements[i]));
            }

            // Define textfield component
            var textFieldComponents = [];
            var TextFieldElements = document.querySelectorAll(".ms-TextField");
            for (i = 0; i < TextFieldElements.length; i++) {
                textFieldComponents.push(new fabric['TextField'](TextFieldElements[i]));
            }

            // Default value for checkbox
            checkBoxComponents[0].check();

            $('#addressTo').val(Office.context.roamingSettings.get('secureEmailAddress'));

            // Get subject from the letter
            item.subject.getAsync(
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.log(asyncResult.error.message);
                    }
                    else {
                        // Successfully got the subject, display it.
                        console.log('The subject is: ' + asyncResult.value);
                        subj = asyncResult.value.trim();
                        emails = subj.match(/(?:#to\s)([^\s]+)/g);
                        if (emails) {
                            emails.forEach(function (email) {
                                $select[0].selectize.addOption({
                                    name: '',
                                    email: email.substr(4)
                                });
                                $select[0].selectize.addItem(email.substr(4));
                                subj = subj.replace(email, '');
                            });
                        } else emails = [];
                        item.to.getAsync(
                            function (asyncResult) {
                                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                    console.log(asyncResult.error.message);
                                }
                                else {
                                    console.log(asyncResult.value);
                                    asyncResult.value.forEach(function (value) {
                                        if (value.emailAddress !== $('#addressTo').val()) {
                                            emails.push('#to ' + value.emailAddress);
                                            $select[0].selectize.addOption({
                                                name: '',
                                                email: value.emailAddress
                                            });
                                            $select[0].selectize.addItem(value.emailAddress);
                                        }
                                    });
                                    console.log(emails);
                                }
                            }
                        );
                        if (subj.search('#allow') + 1) {
                            subj = subj.replace('#allow', '');
                            checkBoxComponents[0].unCheck();
                        }
                        subj = subj.trim().replace(/\s+/g, ' ');
                        $('#subject').val(subj);
                    }
                });
            // Add subject to message
            $('#add').click(function (event) {
                var subj = $('#subject').val();
                console.log(emails);
                Office.context.roamingSettings.set('secureEmailAddress', $('#addressTo').val());
                Office.context.roamingSettings.saveAsync();
                item.to.setAsync([$('#addressTo').val()]);
                $select[0].selectize.items.forEach(function (email) {
                    if (subject !== '') subject = subject + ' ';
                    subject += '#to ' + email;
                });
                if (!checkBoxComponents[0].getValue()) subject = '#allow ' + subject;
                if (subj && subj.length > 0) {
                    subject = subj.trim().replace(/\s+/g, ' ')
                        + ' ' + subject;
                }

                item.subject.setAsync(
                    subject,
                    {asyncContext: {var1: 1, var2: 2}},
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            console.log(asyncResult.error.message);
                        }
                        else {
                            console.log(asyncResult);
                        }
                    });
                console.log(subject);
                subject = '';
                Office.context.ui.closeContainer();
            });
            $('#remove').click(function (event) {
                item.to.setAsync(emails.map(function (email) {
                    return email.substring(4);
                }));
                item.subject.setAsync(
                    subj,
                    {asyncContext: {var1: 1, var2: 2}},
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            console.log(asyncResult.error.message);
                        }
                        else {
                            console.log(asyncResult);
                        }
                    });
                Office.context.ui.closeContainer();
            })
        });
    };

    /**
     * Get contacts from REST API
     * @param accessToken
     * @param callback
     */
    function getContactsItem(accessToken, callback) {
        // Construct the REST URL
        var getContactsUrl = Office.context.mailbox.restUrl +
            '/v2.0/me/contacts?$select=EmailAddresses,GivenName,Surname';

        $.ajax({
            url: getContactsUrl,
            dataType: 'json',
            headers: {'Authorization': 'Bearer ' + accessToken}
        }).done(function (item) {
            var i, j, name;
            for (i = 0; i < item.value.length; i++) {
                if (item.value[i].GivenName || item.value[i].SurName) {
                    name = (item.value[i].GivenName ? item.value[i].GivenName : '') + (item.value[i].Surname ? ' ' + item.value[i].Surname : '');
                } else {
                    name = '<Unknown>';
                }
                for (j = 0; j < item.value[i].EmailAddresses.length; j++) {
                    peopleArray.push({
                        name: name,
                        email: item.value[i].EmailAddresses[j].Address
                    })
                }
            }
            console.log(peopleArray);

            if ($select && peopleArray) {
                peopleArray.forEach(function (people) {
                    $select[0].selectize.addOption({
                        name: people.name,
                        email: people.email
                    });
                });
            }
        }).fail(function (error) {
            // Handle error
            console.log(error);
        });
    }
})();


