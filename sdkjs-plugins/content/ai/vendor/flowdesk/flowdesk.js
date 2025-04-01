(function(window, document) {
    const component = function(config) {
        const referer = 'flow-desk';
        let iframe;

        const add_block = config => {
            post_message(iframe.contentWindow, {
                command: 'addAction',
                referer: 'flow-desk',
                data: config,
            })
        }

        const get_route = () => {
            post_message(iframe.contentWindow, {
                command: 'getRoute',
                referer: referer,
                data: "",
            })
        }

        const on_message = function(msg) {
            let data = msg.data;
            if (Object.prototype.toString.apply(data) !== '[object String]' || !window.JSON) {
                return;
            }
            // var cmd, handler;
            let obj, handler;

            try {
                obj = window.JSON.parse(data)
            } catch(e) {
                obj = undefined;
            }

            if (obj && obj.referer == referer) {
                const events = config.events || {};
                let handler;
                data = {};
                switch (obj.command) {
                    case 'getRoute':
                        handler = events['onGetRoute'];
                        data = obj.data;
                        break;
                }
                if (handler && typeof handler == "function") {
                    const _self = this;
                    handler.call(_self, {target: _self, data: data});
                }
            }
        };

        const parent = document.getElementById(config.boxId);
        if ( parent ) {
            iframe = create_iframe(config);
            if ( config.events?.onLoad )
                iframe.onload = config.events.onLoad;

            parent.appendChild(iframe);

            const msgdispatcher = new MessageDispatcher(on_message, this);
        }

        return {
            addAction: add_block,
            queryRoute: get_route,
        };
    }

    window.FlowDesk = component;

    function getBasePath() {
        const scripts = document.getElementsByTagName('script');
        let match;

        for (let i = scripts.length - 1; i >= 0; i--) {
            match = scripts[i].src.match(/(.*)FlowDesk.js/i);
            if (match) return match[1];
        }

        return "";
    }

    function create_iframe(config) {
        iframe = document.createElement("iframe");
        iframe.width        = '100%';
        iframe.height       = '100%';
        // iframe.align        = "top";
        // iframe.frameBorder  = 0;
        // iframe.scrolling    = "no";

        let _url_ = getBasePath() + 'assets/index.html';
        if ( config ) {
            const params = new URLSearchParams();
            config.miniMap === true && params.append('minimap', "1");
            config.zoomControls === true && params.append('zoomctrls', "1");
            if (params.size)
                _url_ += `?${params.toString()}`;
        }
        iframe.src = _url_;

        return iframe;
    }

    function post_message(wnd, msg) {
        if (wnd && wnd.postMessage && window.JSON) {
            wnd.postMessage(window.JSON.stringify(msg), "*");
        }
    }

    const MessageDispatcher = function(fn, scope) {
        const _fn     = fn,
            _scope  = scope || window,
            eventFn = function(msg) {
                _fn.call(_scope, msg);
            };

        const _bindEvents = function() {
            if (window.addEventListener) {
                window.addEventListener("message", eventFn, false)
            }
            else if (window.attachEvent) {
                window.attachEvent("onmessage", eventFn);
            }
        };

        const _unbindEvents = function() {
            if (window.removeEventListener) {
                window.removeEventListener("message", eventFn, false)
            }
            else if (window.detachEvent) {
                window.detachEvent("onmessage", eventFn);
            }
        };

        _bindEvents.call(this);

        return {
            unbindEvents: _unbindEvents
        }
    };

})(window, document);