Notification.requestPermission();

$(function () {
    // App configuration
    const authEndpoint = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?';
    const redirectUri = 'https://david-codes.hatanian.com/zoom-meeting-notifier';
    const appId = '56a8a6dc-76ab-45de-b45d-91388f4bfbc2';
    const scopes = 'openid profile User.Read Calendars.Read';
    const SEARCH_LOOKAHEAD_HOURS = 20;
    const NOTIFICATION_LOOKAHEAD_MINUTES = 16;
    let ALL_USER_EVENTS = [];

    // Check for browser support for localStorage
    if (typeof(Storage) === 'undefined') {
        render('#unsupportedbrowser');
        return;
    }

    // Check for browser support for crypto.getRandomValues
    const cryptObj = window.crypto || window.msCrypto; // For IE11
    if (cryptObj === undefined || cryptObj.getRandomValues === 'undefined') {
        render('#unsupportedbrowser');
        return;
    }

    render(window.location.hash);

    $(window).on('hashchange', function () {
        render(window.location.hash);
    });

    function render(hash) {

        var action = hash.split('=')[0];

        // Hide everything
        $('.main-container .page').hide();

        // Check for presence of access token
        var isAuthenticated = (localStorage.accessToken != null
                               && localStorage.accessToken.length > 0);
        renderNav(isAuthenticated);

        if (isAuthenticated) {
            setInterval(refreshCalendar, 20000);
            setInterval(notifyUpcomingEvents, 5000);
        }

        function refreshCalendar() {
            ALL_USER_EVENTS = [];
            renderCalendar();
        }

        function notifyUpcomingEvents() {
            if (Notification.permission === "granted") {
                ALL_USER_EVENTS.forEach(notifyUpcomingEvent);
            } else {
                Notification.requestPermission();
            }
        }

        function isSoon(event) {
            let difference = Date.parse(event.start.dateTime) - new Date().getTime();
            return difference >= 0 && difference < NOTIFICATION_LOOKAHEAD_MINUTES * 60 * 1000;
        }

        function notifyUpcomingEvent(event) {
            if (event.zoomLink && isSoon(event)) {
                if (!localStorage[`ZZZalreadyNotified-${event.id}`]) {
                    const options = {
                        body: 'Click to open Zoom',
                        requireInteraction: true,
                        tag: event.id,
                        icon: 'http://localhost:8080/zoom.jpeg'
                    };

                    let notification = new Notification(event.subject, options);
                    notification.onclick = function(clickEvent) {
                        clickEvent.preventDefault();
                        console.log(event.zoomLink);
                        window.open(event.zoomLink);
                    };

                    setTimeout(notification.close.bind(notification), 20 * 60 * 1000);
                    localStorage[`ZZZalreadyNotified-${event.id}`] = true;
                }
            }
        }

        var pagemap = {

            // Welcome page
            '': function () {
                if (isAuthenticated) {
                    renderCalendar();
                } else {
                    // Redirect to home page
                    window.location = buildAuthUrl();
                }
            },

            // Receive access token
            '#access_token': function () {
                handleTokenResponse(hash);
            },

            // Signout
            '#signout': function () {
                clearUserState();

                // Redirect to home page
                window.location.hash = '#';
            },

            // Error display
            '#error': function () {
                var errorresponse = parseHashParams(hash);
                if (errorresponse.error === 'login_required' ||
                    errorresponse.error === 'interaction_required') {
                    // For these errors redirect the browser to the login
                    // page.
                    window.location = buildAuthUrl();
                } else {
                    renderError(errorresponse.error, errorresponse.error_description);
                }
            },

            // Shown if browser doesn't support session storage
            '#unsupportedbrowser': function () {
                $('#unsupported').show();
            }
        };

        if (pagemap[action]) {
            pagemap[action]();
        } else {
            // Redirect to home page
            window.location.hash = '#';
        }
    }

    function setActiveNav(navId) {
        $('#navbar').find('li').removeClass('active');
        $(navId).addClass('active');
    }

    function renderNav(isAuthed) {
        if (isAuthed) {
            $('.authed-nav').show();
            getUserEmailAddress(function (userEmail, error) {
                if (error) {
                    renderError('getUserEmailAddress failed', error.responseText);
                } else {
                    $('#useremail').text(userEmail);
                }
            });
        } else {
            $('.authed-nav').hide();
        }
    }

    function renderError(error, description) {
        $('#error-name', window.parent.document)
            .text('An error occurred: ' + decodePlusEscaped(error));
        $('#error-desc', window.parent.document).text(decodePlusEscaped(description));
        $('#error-display', window.parent.document).show();
    }

    function renderUpdatedEvents() {
        var templateSource = $('#event-list-template').html();
        var template = Handlebars.compile(templateSource);
        var eventList = template({events: ALL_USER_EVENTS});
        $('#event-list').empty();
        $('#event-list').append(eventList);
    }

    function renderCalendar() {
        setActiveNav('#calendar-nav');
        $('#calendar').show();
        // Get user's email address
        getUserEmailAddress(function (userEmail, error) {
            if (error) {
                renderError('getUserEmailAddress failed', error.responseText);
            } else {
                getUserEvents(userEmail, function (events, error) {
                    if (error) {
                        renderError('getUserEvents failed', error);
                    } else {
                        $('#calendar-status').text('Upcoming events:');
                        events.forEach(appendZoomLink);
                        ALL_USER_EVENTS = ALL_USER_EVENTS.concat(events);
                        ALL_USER_EVENTS = ALL_USER_EVENTS.sort(
                            (event1, event2) => event1.start.dateTime.localeCompare(
                                event2.start.dateTime));
                        renderUpdatedEvents();
                    }
                });
            }
        });
    }

    function appendZoomLink(event) {
        zoomLinks = event.body.content.match('https://skyscanner.zoom.us/j/[0-9]+');
        if (zoomLinks != null && zoomLinks.length > 0) {
            event.zoomLink = zoomLinks[0];
        }
    }

    // OAUTH FUNCTIONS =============================

    function buildAuthUrl() {
        // Generate random values for state and nonce
        localStorage.authState = guid();
        localStorage.authNonce = guid();

        var authParams = {
            response_type: 'id_token token',
            client_id: appId,
            redirect_uri: redirectUri,
            scope: scopes,
            state: localStorage.authState,
            nonce: localStorage.authNonce,
            response_mode: 'fragment'
        };

        return authEndpoint + $.param(authParams);
    }

    function handleTokenResponse(hash) {
        // If this was a silent request remove the iframe
        $('#auth-iframe').remove();

        // clear tokens
        localStorage.removeItem('accessToken');
        localStorage.removeItem('idToken');

        var tokenresponse = parseHashParams(hash);

        // Check that state is what we sent in sign in request
        if (tokenresponse.state != localStorage.authState) {
            localStorage.removeItem('authState');
            localStorage.removeItem('authNonce');
            // Report error
            window.location.hash =
                '#error=Invalid+state&error_description=The+state+in+the+authorization+response+did+not+match+the+expected+value.+Please+try+signing+in+again.';
            return;
        }

        localStorage.authState = '';
        localStorage.accessToken = tokenresponse.access_token;

        // Get the number of seconds the token is valid for,
        // Subract 5 minutes (300 sec) to account for differences in clock settings
        // Convert to milliseconds
        var expiresin = (parseInt(tokenresponse.expires_in) - 300) * 1000;
        var now = new Date();
        var expireDate = new Date(now.getTime() + expiresin);
        localStorage.tokenExpires = expireDate.getTime();

        localStorage.idToken = tokenresponse.id_token;

        validateIdToken(function (isValid) {
            if (isValid) {
                // Redirect to home page
                window.location.hash = '#';
            } else {
                clearUserState();
                // Report error
                window.location.hash =
                    '#error=Invalid+ID+token&error_description=ID+token+failed+validation,+please+try+signing+in+again.';
            }
        });
    }

    function validateIdToken(callback) {
        // Per Azure docs (and OpenID spec), we MUST validate
        // the ID token before using it. However, full validation
        // of the signature currently requires a server-side component
        // to fetch the public signing keys from Azure. This sample will
        // skip that part (technically violating the OpenID spec) and do
        // minimal validation

        if (null == localStorage.idToken || localStorage.idToken.length <= 0) {
            callback(false);
        }

        // JWT is in three parts seperated by '.'
        var tokenParts = localStorage.idToken.split('.');
        if (tokenParts.length != 3) {
            callback(false);
        }

        // Parse the token parts
        var header = KJUR.jws.JWS.readSafeJSONString(b64utoutf8(tokenParts[0]));
        var payload = KJUR.jws.JWS.readSafeJSONString(b64utoutf8(tokenParts[1]));

        // Check the nonce
        if (payload.nonce != localStorage.authNonce) {
            localStorage.authNonce = '';
            callback(false);
        }

        localStorage.authNonce = '';

        // Check the audience
        if (payload.aud != appId) {
            callback(false);
        }

        // Check the issuer
        // Should be https://login.microsoftonline.com/{tenantid}/v2.0
        if (payload.iss !== 'https://login.microsoftonline.com/' + payload.tid + '/v2.0') {
            callback(false);
        }

        // Check the valid dates
        var now = new Date();
        // To allow for slight inconsistencies in system clocks, adjust by 5 minutes
        var notBefore = new Date((payload.nbf - 300) * 1000);
        var expires = new Date((payload.exp + 300) * 1000);
        if (now < notBefore || now > expires) {
            callback(false);
        }

        // Now that we've passed our checks, save the bits of data
        // we need from the token.

        localStorage.userDisplayName = payload.name;
        localStorage.userSigninName = payload.preferred_username;

        // Per the docs at:
        // https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-protocols-implicit/#send-the-sign-in-request
        // Check if this is a consumer account so we can set domain_hint properly
        localStorage.userDomainType =
            payload.tid === '9188040d-6c67-4c5b-b112-36a304b66dad' ? 'consumers' : 'organizations';

        callback(true);
    }

    function makeSilentTokenRequest(callback) {
        // Build up a hidden iframe
        var iframe = $('<iframe/>');
        iframe.attr('id', 'auth-iframe');
        iframe.attr('name', 'auth-iframe');
        iframe.appendTo('body');
        iframe.hide();

        iframe.load(function () {
            callback(localStorage.accessToken);
        });

        iframe.attr('src', buildAuthUrl() + '&prompt=none&domain_hint=' +
                           localStorage.userDomainType + '&login_hint=' +
                           localStorage.userSigninName);
    }

    // Helper method to validate token and refresh
    // if needed
    function getAccessToken(callback) {
        var now = new Date().getTime();
        var isExpired = now > parseInt(localStorage.tokenExpires);
        // Do we have a token already?
        if (localStorage.accessToken && !isExpired) {
            // Just return what we have
            if (callback) {
                callback(localStorage.accessToken);
            }
        } else {
            // Attempt to do a hidden iframe request
            makeSilentTokenRequest(callback);
        }
    }

    // OUTLOOK API FUNCTIONS =======================
    function getUserEmailAddress(callback) {
        if (localStorage.userEmail) {
            return localStorage.userEmail;
        } else {
            getAccessToken(function (accessToken) {
                if (accessToken) {
                    // Create a Graph client
                    var client = MicrosoftGraph.Client.init({
                                                                authProvider: (done) => {
                                                                    // Just return the token
                                                                    done(null, accessToken);
                                                                }
                                                            });

                    // Get the Graph /Me endpoint to get user email address
                    client
                        .api('/me')
                        .get((err, res) => {
                            if (err) {
                                callback(null, err);
                            } else {
                                callback(res.mail);
                            }
                        });
                } else {
                    var error = {responseText: 'Could not retrieve access token'};
                    callback(null, error);
                }
            });
        }
    }

    function getUserEvents(emailAddress, callback) {
        getAccessToken(function (accessToken) {
            if (accessToken) {
                // Create a Graph client
                var client = MicrosoftGraph.Client.init({
                                                            authProvider: (done) => {
                                                                done(null, accessToken);
                                                            }
                                                        });

                client
                    .api('/me/calendars')
                    .get((err, res) => {
                        if (err) {
                            callback(null, err);
                        } else {
                            res.value.forEach(
                                calendar => getCalendarEventsForCalendar(calendar.id, emailAddress,
                                                                         client, callback)
                            );
                        }
                    });
            } else {
                const error = {responseText: 'Could not retrieve access token'};
                callback(null, error);
            }
        });
    }

    function getCalendarEventsForCalendar(calendarId, emailAddress, client, callback) {
        const startDateTime = new Date().toISOString();
        let inFuture = new Date();
        inFuture.setHours(inFuture.getHours() + SEARCH_LOOKAHEAD_HOURS);
        const endDateTime = inFuture.toISOString();

        const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
        client
            .api(`/me/calendars/${calendarId}/calendarview`
                 + `?startDateTime=${startDateTime}`
                 + `&endDateTime=${endDateTime}`)
            .header('X-AnchorMailbox', emailAddress)
            .header('Prefer', `outlook.timezone="${timeZone}"`)
            .get((err, res) => {
                if (err) {
                    callback(null, err);
                } else {
                    callback(res.value);
                }
            });
    }

    // HELPER FUNCTIONS ============================

    function guid() {
        var buf = new Uint16Array(8);
        cryptObj.getRandomValues(buf);
        function s4(num) {
            var ret = num.toString(16);
            while (ret.length < 4) {
                ret = '0' + ret;
            }
            return ret;
        }

        return s4(buf[0]) + s4(buf[1]) + '-' + s4(buf[2]) + '-' + s4(buf[3]) + '-' +
               s4(buf[4]) + '-' + s4(buf[5]) + s4(buf[6]) + s4(buf[7]);
    }

    function parseHashParams(hash) {
        var params = hash.slice(1).split('&');

        var paramarray = {};
        params.forEach(function (param) {
            param = param.split('=');
            paramarray[param[0]] = param[1];
        });

        return paramarray;
    }

    function decodePlusEscaped(value) {
        // decodeURIComponent doesn't handle spaces escaped
        // as '+'
        if (value) {
            return decodeURIComponent(value.replace(/\+/g, ' '));
        } else {
            return '';
        }
    }

    function clearUserState() {
        // Clear session
        localStorage.clear();
    }

    Handlebars.registerHelper("formatDate", function (datetime) {
        // Dates from API look like:
        // 2016-06-27T14:06:13Z

        var date = new Date(datetime);
        return date.toLocaleDateString() + ' ' + date.toLocaleTimeString();
    });
});
