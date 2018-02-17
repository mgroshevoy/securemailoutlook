'use strict';

(function () {
    var item, $select, $selectcc, $selectbcc, subject = '', subj, peopleArray = [];

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        item = Office.context.mailbox.item;
        item.saveAsync(function (result) {
            // Get token from client
            Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function (result) {
                if (result.status === "succeeded") {
                    var accessToken = result.value;
                    getContactsItem(accessToken);
                } else {
                    // Handle the error
                    console.error(result);
                    getTokenFromClient();
                }
            });
        });

        // Invoke by Contoso Subject and CC Checker add-in before send is allowed.
        // <param name="event">ItemSend event is automatically passed by on send code to the function specified in the manifest.</param>

        $(document).ready(function () {


            var CalloutExamples = document.querySelectorAll(".ms-CalloutExample");
            for (var i = 0; i < CalloutExamples.length; i++) {
                var Example = CalloutExamples[i];
                var ExampleButtonElement = Example.querySelector(".ms-CalloutExample-button");
                var CalloutElement = Example.querySelector(".ms-Callout");
                new fabric['Callout'](
                    CalloutElement,
                    ExampleButtonElement,
                    "right"
                );
            }

            var REGEX_EMAIL = '(([^<>()\\[\\]\\\\.,;:\\s@"]+(\\.[^<>()\\[\\]\\\\.,;:\\s@"]+)*)|(".+"))@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}])|(([a-zA-Z\\-0-9]+\\.)+[a-zA-Z]{2,}))';

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
                    regex = new RegExp('^' + REGEX_EMAIL + '$', 'ig');
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

            $selectcc = $('#select-cc').selectize({
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
                    regex = new RegExp('^' + REGEX_EMAIL + '$', 'ig');
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

            $selectbcc = $('#select-bcc').selectize({
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
                    regex = new RegExp('^' + REGEX_EMAIL + '$', 'ig');
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
            for (i = 0; i < ButtonElements.length; i++) {
                new fabric['Button'](ButtonElements[i], function (event) {
                });
            }

            // Define checkbox component
            var checkBoxComponents = [];
            var CheckBoxElements = document.querySelectorAll(".ms-CheckBox");
            for (i = 0; i < CheckBoxElements.length; i++) {
                checkBoxComponents.push(new fabric['CheckBox'](CheckBoxElements[i]));
            }

            // Define textfield component
            var textFieldComponents = [];
            var TextFieldElements = document.querySelectorAll(".ms-TextField");
            for (i = 0; i < TextFieldElements.length; i++) {
                textFieldComponents.push(new fabric['TextField'](TextFieldElements[i]));
            }

            textFieldComponents[0]._textField.addEventListener('change', function (e) {
                console.log(e);
            });

            // Default value for checkbox
            checkBoxComponents[0].check();

            $('#addressTo').val(Office.context.roamingSettings.get('secureEmailAddress'));
            if ($('#addressTo').val().match(new RegExp('^' + REGEX_EMAIL + '$', 'ig'))) {
                $('#action-area').show();
            } else {
                $('#action-area').hide();
            }

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

                        item.to.getAsync(
                            function (asyncResult) {
                                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                    console.log(asyncResult.error.message);
                                }
                                else {
                                    asyncResult.value.forEach(function (value) {
                                        if (value.emailAddress !== $('#addressTo').val()) {
                                            $select[0].selectize.addOption({
                                                name: value.displayName,
                                                email: value.emailAddress
                                            });
                                            $select[0].selectize.addItem(value.emailAddress);
                                        }
                                    });
                                }
                            }
                        );

                        item.cc.getAsync(
                            function (asyncResult) {
                                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                    console.log(asyncResult.error.message);
                                }
                                else {
                                    asyncResult.value.forEach(function (value) {
                                        if (value.emailAddress !== $('#addressTo').val()) {
                                            $selectcc[0].selectize.addOption({
                                                name: value.displayName,
                                                email: value.emailAddress
                                            });
                                            $selectcc[0].selectize.addItem(value.emailAddress);
                                        }
                                    });
                                }
                            }
                        );

                        item.bcc.getAsync(
                            function (asyncResult) {
                                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                    console.log(asyncResult.error.message);
                                }
                                else {
                                    asyncResult.value.forEach(function (value) {
                                        if (value.emailAddress !== $('#addressTo').val()) {
                                            $selectbcc[0].selectize.addOption({
                                                name: value.displayName,
                                                email: value.emailAddress
                                            });
                                            $selectbcc[0].selectize.addItem(value.emailAddress);
                                        }
                                    });
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

            $('#add').click(function () {
                var subj = $('#subject').val();
                var arrayTo = [];
                var secureEmailAddress = $('#addressTo').val();
                var secureDomain = secureEmailAddress.match(/@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))/);
                if (secureDomain) secureDomain = secureDomain[0];

                arrayTo = [];
                $select[0].selectize.items.forEach(function (item) {
                    var name = '';
                    $.each($select[0].selectize.options, function (code, recipient) {
                        if (recipient.email === item && recipient.name && recipient.name != item) {
                            name = recipient.name;
                        }
                    });
                    if (item.match(secureDomain)) {
                        arrayTo.push({displayName: name, emailAddress: item});
                    } else {
                        arrayTo.push({displayName: name, emailAddress: item.replace('@', '.at.') + secureDomain});
                    }
                });

                item.to.setAsync(arrayTo);
                arrayTo = [];
                $selectcc[0].selectize.items.forEach(function (item) {
                    var name = '';
                    $.each($selectcc[0].selectize.options, function (code, recipient) {
                        if (recipient.email === item && recipient.name && recipient.name != item) {
                            name = recipient.name;
                        }
                    });
                    if (item.match(secureDomain)) {
                        arrayTo.push({displayName: name, emailAddress: item});
                    } else {
                        arrayTo.push({displayName: name, emailAddress: item.replace('@', '.at.') + secureDomain});
                    }
                });
                if (arrayTo.length) item.cc.setAsync(arrayTo);
                arrayTo = [];
                $selectbcc[0].selectize.items.forEach(function (item) {
                    var name = '';
                    $.each($selectbcc[0].selectize.options, function (code, recipient) {
                        if (recipient.email === item && recipient.name && recipient.name != item) {
                            name = recipient.name;
                        }
                    });
                    if (item.match(secureDomain)) {
                        arrayTo.push({displayName: name, emailAddress: item});
                    } else {
                        arrayTo.push({displayName: name, emailAddress: item.replace('@', '.at.') + secureDomain});
                    }
                });
                if (arrayTo.length) item.bcc.setAsync(arrayTo);

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

            // if addressTo empty then hide the main section

            $('#addressTo').on('input', function () {

                if ($(this).val().match(new RegExp('^' + REGEX_EMAIL + '$', 'ig'))) {
                    $('#action-area').show();
                    var secureEmailAddress = $('#addressTo').val();
                    Office.context.roamingSettings.set('secureEmailAddress', secureEmailAddress);
                    Office.context.roamingSettings.saveAsync();
                    Office.context.mailbox.item.notificationMessages.removeAsync("syncError");
                    console.log(secureEmailAddress);
                } else {
                    $('#action-area').hide();
                    Office.context.mailbox.item.notificationMessages.addAsync("syncError", {
                        type: "errorMessage",
                        message: "Please set your secure email address in the SecureMail add-in!"
                    });
                }
            });


            $('#remove').click(function () {
                var secureEmailAddress = $('#addressTo').val();
                var secureDomain = secureEmailAddress.match(/@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?/);

                item.to.setAsync($select[0].selectize.items.map(function (item) {
                    var name = '';
                    $.each($select[0].selectize.options, function (code, recipient) {
                        if (recipient.email === item && recipient.name && recipient.name != item) {
                            name = recipient.name;
                        }
                    });
                    if (item.match(/\.at\./) && item.match(secureDomain)) {
                        return {
                            displayName: name,
                            emailAddress: item.replace(secureDomain, '').replace(/^(.*)\.at\.(.*?)$/, '$1@$2')
                        };
                    } else {
                        return {displayName: name, emailAddress: item};
                    }
                }));

                if ($selectcc[0].selectize.items.length) {
                    item.cc.setAsync($selectcc[0].selectize.items.map(function (item) {
                        var name = '';
                        $.each($selectcc[0].selectize.options, function (code, recipient) {
                            if (recipient.email === item && recipient.name && recipient.name != item) {
                                name = recipient.name;
                            }
                        });
                        if (item.match(/\.at\./) && item.match(secureDomain)) {
                            return {
                                displayName: name,
                                emailAddress: item.replace(secureDomain, '').replace(/^(.*)\.at\.(.*?)$/, '$1@$2')
                            };
                        } else {
                            return {displayName: name, emailAddress: item};
                        }
                    }));
                }

                if ($selectbcc[0].selectize.items.length) {
                    item.bcc.setAsync($selectbcc[0].selectize.items.map(function (item) {
                        var name = '';
                        $.each($selectbcc[0].selectize.options, function (code, recipient) {
                            if (recipient.email === item && recipient.name && recipient.name != item) {
                                name = recipient.name;
                            }
                        });
                        if (item.match(/\.at\./) && item.match(secureDomain)) {
                            return {
                                displayName: name,
                                emailAddress: item.replace(secureDomain, '').replace(/^(.*)\.at\.(.*?)$/, '$1@$2')
                            };
                        } else {
                            return {displayName: name, emailAddress: item};
                        }
                    }));
                }

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
                    name = '';
                }
                for (j = 0; j < item.value[i].EmailAddresses.length; j++) {
                    peopleArray.push({
                        name: name,
                        email: item.value[i].EmailAddresses[j].Address
                    })
                }
            }
            //console.log(peopleArray);

            if ($select && peopleArray) {
                peopleArray.forEach(function (people) {
                    $select[0].selectize.addOption({
                        name: people.name === people.email ? '' : people.name,
                        email: people.email
                    });
                });
            }

            if ($selectcc && peopleArray) {
                peopleArray.forEach(function (people) {
                    $selectcc[0].selectize.addOption({
                        name: people.name === people.email ? '' : people.name,
                        email: people.email
                    });
                });
            }

            if ($selectbcc && peopleArray) {
                peopleArray.forEach(function (people) {
                    $selectbcc[0].selectize.addOption({
                        name: people.name === people.email ? '' : people.name,
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


