(function () {
    var config = {
        clientId: "e9a57c07-69e4-41f5-b868-0f123d2fda17",
        redirectUri: window.location.origin,
        interactionMode: "popUp",
        scopes: ["user.read", "files.read.all", "sites.read.all"],
        authority: "https://login.microsoftonline.com/common"
    }

    var clientApplication = createMsalApplication(config);
    var baseUrl = getQueryVariable("baseUrl")
    window.msGraphApiRoot = (baseUrl) ? baseUrl : "https://graph.microsoft.com/v1.0/me";

    // Wire up to the commands in the html
    var $signInButton = $("#od-login");
    var $signOutButton = $("#od-logoff");
    var $title = $("#od-title");
    var $loading = $("#od-loading");
    var $breadcrumb = $("#od-breadcrumb");
    var $content = $("#od-content");
    var $items = $("#od-items");
    var $json = $("#od-json");

    window.onhashchange = function () {
        loadView(stripHash(window.location.hash));
    }

    window.onload = function () {
        updateSignedInUserState(null);
    }

    $signOutButton.click(function () {
        clientApplication.clearCache();
        clientApplication.user = null;
        updateSignedInUserState();
    });

    $signInButton.click(function () {
        clientApplication.loginPopup(config.scopes).then(function (idToken) {
            // Update the UI status now that we're signed in.
            updateSignedInUserState(true);
        });
    });
    
    // we bind to jquery's ajax start/stop events so that we can style the
    // page differently when a network call is being made
    $(document).on({
        ajaxStart: function() {$('body').addClass('loading');},
        ajaxStop:  function() {$('body').removeClass('loading');}
    });    

    function createMsalApplication(config, authCallback) {
        var msalApp = new Msal.UserAgentApplication(config.clientId, null /*config.authority*/, authCallback);
        msalApp.redirectUri = config.redirectUri;
        msalApp.interactionMode = config.interactionMode;

        var isCallback = msalApp.isCallback(window.location.hash);
        if (isCallback)
        {
            msalApp.handleAuthenticationResponse(window.location.hash);
        }
        return msalApp;
    }

    function updateSignedInUserState(reloadView) {
        // Check login status, update the UI
        var user = clientApplication.getUser();
        if (user) {
            // signed in
            $signInButton.hide();
            $signOutButton.show();
            saveToCookie( { "apiRoot": window.msGraphApiRoot, "signedin": true } );
            if (reloadView) {
                $(window).trigger("hashchange");
            }
        } else {
            // signed out
            $signInButton.show();
            $signOutButton.hide();
            saveToCookie( { "apiRoot": window.msGraphApiRoot, "signedin": false } );
            $breadcrumb.empty();
            $items.empty();
            $json.empty();
        }
    }

    function getUrlParts(url)
    {
        var a = document.createElement("a");
        a.href = url;

        return { "hostname": a.hostname,
                "path": a.pathname }
    }

    function stripHash(view) {
        return view.substr(view.indexOf('#') + 1);
    }

    function saveToCookie(obj) {
        var expiration = new Date();
        expiration.setTime(expiration.getTime() + 3600 * 1000);
        var data = JSON.stringify(obj);
        var cookie = "odexplorer=" + data +"; path=/; expires=" + expiration.toUTCString();

        if (document.location.protocol.toLowerCase() == "https") {
        cookie = cookie + ";secure";
        }
        document.cookie = cookie;
    }

    function loadFromCookie() {
        var cookies = document.cookie;
        var name = "odexplorer=";
        var start = cookies.indexOf(name);
        if (start >= 0) {
            start += name.length;
            var end = cookies.indexOf(';', start);
            if (end < 0) {
                end = cookies.length;
            } else {
                postCookie = cookies.substring(end);
            }

            var value = cookies.substring(start, end);
            return JSON.parse(value);
        }

        return "";
    }

    function getQueryVariable(variable) {
        var query = window.location.search.substring(1);
        var vars = query.split("&");
        for (var i=0;i<vars.length;i++) {
            var pair = vars[i].split("=");
            if(pair[0] == variable) {
                return pair[1];
            }
        }
        return(false);
    }    
    
    function loadView(path) {
        var user = clientApplication.getUser();
        if (!user) {
            return;
        }
        var data = loadFromCookie();
        if (data)
        {
            if (!baseUrl) {
                window.msGraphApiRoot = data.apiRoot;
            }
        }

        // we extract the onedrive path from the url fragment and we
        // flank it with colons to use the api's path-based addressing scheme
        var beforePath = "";
        var afterPath = "";
        if (path.length > 0) {
            beforePath =":";
            afterPath = ":";
        }

        var odurl = msGraphApiRoot + "/drive/root" + beforePath + path + afterPath;

        // the expand and select parameters mean:
        //  "for the item i'm addressing, include its thumbnails and children,
        //   and for each of the children, include its thumbnails. for those
        //   thumbnails, return the 'large' size"
        var thumbnailSize = "large"
        var odquery = "?expand=thumbnails,children(expand=thumbnails(select=" + thumbnailSize + "))";

        getAccessToken(scopes, function(token, error) {
            if (token) {
                loadDriveItemChildren(odurl, odquery, token, path, thumbnailSize);
            } else {
                alert(error);
            }
        })
    }
    
    function loadDriveItemChildren(odurl, odquery, token, path, thumbnailSize) {
        $.ajax({
            url: odurl + odquery,
            dataType: 'json',
            headers: { "Authorization": "Bearer " + token },
            accept: "application/json;odata.metadata=none",
            success: function (data) {
                if (data) {
                    // clear out the old content
                    $items.empty();
                    $json.empty();

                    // add the syntax-highlighted json response
                    $("<pre>").html(syntaxHighlight(data)).appendTo($json);

                    // process the response data. if we get back children (data.children)
                    // then render the tile view. otherwise, render the "one-up" view
                    // for the item's individual data. we also look for children in
                    // 'data.value' because if this app is ever configured to reqeust
                    // '/children' directly instead of '/parent?expand=children', then
                    // they'll be in an array called 'data'
                    var decodedPath = decodeURIComponent(path);
                    document.title = "OneDrive Explorer" + ((decodedPath.length > 0) ? " - " + decodedPath : "");
                        
                    updateBreadcrumb(decodedPath);
                    var children = data.children || data.value;
                    if (children && children.length > 0) {
                        $.each(children, function(i,item) {
                            var tile = $("<div>").
                                attr("href", "#" + path + "/" + encodeURIComponent(item.name)).
                                addClass("item").
                                click(function() {
                                // when the page changes in response to a user click,
                                // we set loadedForHash to the new value and call
                                // odauth ourselves in user-click mode. this causes
                                // the catch-all hashchange event handler not to
                                // process the page again. see comment at the top.
                                loadedForHash = $(this).attr('href');
                                window.location = loadedForHash;
                                }).
                                appendTo($items);

                            // look for various facets on the items and style them accordingly
                            if (item.folder) {
                                tile.addClass("folder");
                            }
                            if (item.file) {
                                tile.addClass("file");
                            }

                            if (item.thumbnails && item.thumbnails.length > 0) {
                                var container = $("<div>").attr("class", "img-container").appendTo(tile)
                                $("<img>").
                                attr("src", item.thumbnails[0][thumbnailSize].url).
                                appendTo(container);
                            }

                            $("<div>").
                                addClass("nameplate").
                                text(item.name).
                                appendTo(tile);
                        });
                    } else if (data.file) {
                        // 1-up view
                        var tile = $("<div>").
                        addClass("item").
                        addClass("oneup").
                        appendTo($items);

                        var downloadUrl = data['@microsoft.graph.downloadUrl'];
                        if (downloadUrl) {
                            tile.click(function() {
                                window.open(downloadUrl, "Download");
                            });
                        }

                        if (data.folder) {
                            tile.addClass("folder");
                        }

                        if (data.thumbnails && data.thumbnails.length > 0) {
                            $("<img>").
                            attr("src", data.thumbnails[0].large.url).
                            appendTo(tile);
                        }
                    } else {
                        $('<p>No items in this folder.</p>').appendTo('#od-items');  
                    }
                } else {
                    // No data was received
                    $items.empty();
                    $json.empty();
                    $('<p>error.</p>').appendTo($items);
                }
            }
        });
    }


    // based on http://jsfiddle.net/KJQ9K/554/
    function syntaxHighlight(obj) {
        var json = JSON.stringify(obj, undefined, 2)
        json = json.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return json.replace(/("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g,
        function (match) {
            var cls = 'number';
            if (/^"/.test(match)) {
            if (/:$/.test(match)) {
                cls = 'key';
            } else {
                cls = 'string';
            }
            } else if (/true|false/.test(match)) {
            cls = 'boolean';
            } else if (/null/.test(match)) {
            cls = 'null';
            }
            return '<span class="' + cls + '">' + match + '</span>';
        });
    }

    // called to update the breadcrumb bar at the top of the page
    function updateBreadcrumb(decodedPath) {
        var path = decodedPath || '';
        $breadcrumb.empty();
        var runningPath = '';
        var segments = path.split('/');
        for (var i = 0 ; i < segments.length; i++) {
        if (i > 0) {
            $('<span>').text(' > ').appendTo($breadcrumb);
        }

        var segment = segments[i];
        if (segment) {
            runningPath = runningPath + '/' + encodeURIComponent(segment);
        } else {
            segment = 'Files';
        }

        $('<a>').
            attr("href", "#" + runningPath).
            click(function() {
                // when the page changes in response to a user click,
                // we set loadedForHash to the new value and call
                // odauth ourselves in user-click mode. this causes
                // the catch-all hashchange event handler not to
                // process the page again. see comment at the top.
                loadedForHash = $(this).attr('href');
                window.location = loadedForHash;
            }).text(segment).appendTo($breadcrumb);
        }
    }    
}());