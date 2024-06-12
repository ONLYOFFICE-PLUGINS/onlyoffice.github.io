/* eslint-disable no-invalid-this */
/* eslint-disable max-len */
/* eslint-disable indent */
'use strict';
/**
 * @typedef {Object} Color
 * @property {String} mainColor
 * @property {String} disabledColor
 */
/**
 * @typedef {Object} Settings
 * @property {Color} backgroundColor
 * @property {Color} progressColor
 * @property {Color} borderProgressColor
 * @property {Color} thumbColor
 * @property {Color} borderBackgroundColor
 * @property {Number} borderRadius
 * @property {Number} thumbRadius
 * @property {Number} height
 * @property {Number} max
 * @property {Number} min
 * @property {Number} startValue
 * @property {Boolean} isEnabledByDefault
 * @property {Boolean} isVertical
 */
/**
 * Slider module
 */
// eslint-disable-next-line no-unused-vars
const SliderModule = (function() {
    /**
     * Default settings for slider
     * @type {Settings}
     */
    const defaultSliderSettings = {
        backgroundColor: {mainColor: '#c0c0c0', disabledColor: '#c0c0c0'},
        progressColor: {mainColor: '#444444', disabledColor: '#a0a0a0'},
        borderProgressColor: {mainColor: '#444444', disabledColor: '#a0a0a0'},
        thumbColor: {mainColor: '#444444', disabledColor: '#a0a0a0'},
        borderBackgroundColor: {mainColor: '#c0c0c0', disabledColor: '#c0c0c0'},
        borderRadius: 2,
        thumbRadius: 6,
        height: 4,
        max: 100,
        min: -100,
        startValue: 0,
        isEnabledByDefault: false,
        isVertical: false,
    };
    /**
     * @param {Number} x
     * @param {Number} y
     */
    function Coordinates(x, y) {
        this.x = x;
        this.y = y;
    }
    return {
        changeDefaultSettings: function(settings) {
            // eslint-disable-next-line prefer-const
            for (let i in settings) {
                if (Object.prototype.hasOwnProperty.call(defaultSliderSettings, i)) {
                    defaultSliderSettings[i] = settings[i] || defaultSliderSettings[i];
                }
            }
        },
        /**
         * Constructor for sliders
         * @param {Settings} oSettings
         * @param {HTMLDivElement} parent
         */
        CSlider: function(oSettings, parent) {
            // TODO: refactoring
            const settings = {};
            this.handlers = {};
            this.parent = parent;
            // eslint-disable-next-line prefer-const, guard-for-in
            for (let i in defaultSliderSettings) {
                settings[i] = oSettings[i] || defaultSliderSettings[i];
            }
            let isEnabled = settings.isEnabledByDefault;
            if (settings.borderRadius > settings.height / 2) {
                settings.borderRadius = settings.height / 2;
            }
            if (settings.thumbRadius * 2 < settings.height) {
                settings.thumbRadius = settings.height / 2;
            }
            if (parent.hasAttribute('min')) {
                settings.min = +parent.getAttribute('min');
            }
            if (parent.hasAttribute('max')) {
                settings.max = +parent.getAttribute('max');
            }
            let value = settings.startValue;
            this.parent.setAttribute('value', '' + Math.floor(value));

            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d');
            /**
             */
            function update() {
                const cWidth = (settings.isVertical) ? canvas.height : canvas.width;
                const cHeight = (settings.isVertical) ? canvas.width : canvas.height;
                const borderRadius = settings.borderRadius * window.devicePixelRatio;
                const height = settings.height * window.devicePixelRatio;
                const thumbRadius = settings.thumbRadius * window.devicePixelRatio;
                const total = settings.max - settings.min;
                const percent = (value - settings.min) / total;
                const needx = ((cWidth - 2 * thumbRadius) * percent) + thumbRadius;

                context.clearRect(0, 0, canvas.width, canvas.height);

                const leftTop = new Coordinates((borderRadius * 2) >> 0, (cHeight / 2 - height / 2) >> 0);
                const leftMid = new Coordinates(0, (cHeight / 2) >> 0);
                const leftBot = new Coordinates((borderRadius * 2) >> 0, (cHeight / 2 + height / 2) >> 0);

                const centerTop = new Coordinates(needx >> 0, (cHeight / 2 - height / 2) >> 0);
                const centerBot = new Coordinates(needx >> 0, (cHeight / 2 + height / 2) >> 0);

                const rightTop = new Coordinates((cWidth - 2 * borderRadius) >> 0, (cHeight / 2 - height / 2) >> 0);
                const rightMid = new Coordinates(cWidth >> 0, (cHeight / 2) >> 0);
                const rightBot = new Coordinates((cWidth - 2 * borderRadius) >> 0, (cHeight / 2 + height / 2) >> 0);

                const thumbCoords = new Coordinates(needx >> 0, (cHeight / 2) >> 0);

                const needPenW = 1;
                context.lineWidth = (needPenW * window.devicePixelRatio) >> 0;
                const offsetCeil = (context.lineWidth & 1) / 2;

                [leftTop, leftMid, leftBot, centerTop, centerBot, rightTop, rightMid, rightBot, thumbCoords].forEach(function(item) {
                    if (settings.isVertical) {
                        item.y = [(cWidth - item.x), item.x = item.y][0];
                        item.x += offsetCeil;
                    } else {
                        item.y += offsetCeil;
                    }
                });

                // before thumb
                context.beginPath();
                context.moveTo(leftTop.x, leftTop.y);
                if (settings.isVertical) {
                    context.arcTo(leftTop.x, (leftTop.y) + borderRadius, leftMid.x, (leftBot.y) + borderRadius, borderRadius);
                    context.arcTo(leftBot.x, (leftBot.y) + borderRadius, (leftBot.x), (leftBot.y), borderRadius);
                } else {
                    context.arcTo((leftTop.x) - borderRadius, leftTop.y, (leftBot.x) - borderRadius, leftMid.y, borderRadius);
                    context.arcTo((leftBot.x) - borderRadius, leftBot.y, (leftBot.x), leftBot.y, borderRadius);
                }
                context.lineTo(centerBot.x, centerBot.y);
                context.lineTo(centerTop.x, centerTop.y);
                context.closePath();
                context.fillStyle = (isEnabled) ? settings.progressColor.mainColor : settings.progressColor.disabledColor;
                context.fill();
                context.strokeStyle = (isEnabled) ? settings.borderProgressColor.mainColor : settings.borderProgressColor.disabledColor;
                context.stroke();
                // after thumb
                context.beginPath();
                context.moveTo(rightTop.x, (rightTop.y));
                if (settings.isVertical) {
                    context.arcTo(rightTop.x, (rightTop.y) - borderRadius, rightMid.x, (rightBot.y) - borderRadius, borderRadius);
                    context.arcTo(rightBot.x, (rightBot.y) - borderRadius, rightBot.x, (rightBot.y), borderRadius);
                } else {
                    context.arcTo((rightTop.x) + borderRadius, rightTop.y, (rightBot.x) + borderRadius, rightMid.y, borderRadius);
                    context.arcTo((rightBot.x) + borderRadius, rightBot.y, rightBot.x, rightBot.y, borderRadius);
                }
                context.lineTo(centerBot.x, centerBot.y);
                context.lineTo(centerTop.x, centerTop.y);
                context.closePath();
                context.fillStyle = (isEnabled) ? settings.backgroundColor.mainColor : settings.backgroundColor.disabledColor;
                context.fill();
                context.strokeStyle = (isEnabled) ? settings.borderBackgroundColor.mainColor : settings.borderBackgroundColor.disabledColor;
                context.stroke();
                // Thumb
                context.beginPath();
                context.arc(thumbCoords.x, thumbCoords.y, thumbRadius, 0, 2*Math.PI);
                context.fillStyle = (isEnabled) ? settings.thumbColor.mainColor : settings.thumbColor.disabledColor;
                context.fill();
            }
            const sendEvent = _sendEvent.bind(this);
            canvas.onpointerdown = function(event) {
                if (isEnabled) {
                    event.preventDefault();
                    event.stopPropagation();
                    const pageX = (settings.isVertical) ? event.pageY : event.pageX;
                    moveAt(pageX);
                    /**
                     * @param {Number} pageX
                     */
                    function moveAt(pageX) {
                        const rect = parent.getBoundingClientRect();
                        const offset = (settings.isVertical) ? rect.top : rect.left;
                        const oWidth = (settings.isVertical) ? rect.height : rect.width;
                        const pageOffset = (settings.isVertical) ? window.pageYOffset : window.pageXOffset;
                        const oValue = value;
                        let moving = (pageX - offset - settings.thumbRadius - pageOffset);
                        if (settings.isVertical) {
                            moving = oWidth - moving - 2*settings.thumbRadius;
                        }
                        if (moving < 0) {
                            value = settings.min;
                        } else if (moving > oWidth - settings.thumbRadius * 2) {
                            value = settings.max;
                        } else {
                            const total = settings.max - settings.min;
                            const maxWidth = oWidth - settings.thumbRadius * 2;
                            const percent = moving / maxWidth;
                            value = (settings.min + percent * total);
                            // value = Math.floor(value);
                        }
                        if (Math.floor(oValue) != Math.floor(value)) {
                            parent.setAttribute('value', '' + Math.floor(value));
                            update();
                            sendEvent('change');
                        }
                    }
                    /**
                     * @param {PointerEvent} event
                     */
                    function onPointerMove(event) {
                        const pageX = (settings.isVertical) ? event.pageY : event.pageX;
                        moveAt(pageX);
                    }
                    /**
                     * @param {PointerEvent} event
                     */
                    function onPointerUp(event) {
                        document.removeEventListener('pointermove', onPointerMove);
                        document.removeEventListener('pointerup', onPointerUp);
                    }
                    document.addEventListener('pointermove', onPointerMove, false);
                    document.addEventListener('pointerup', onPointerUp, false);
                }
            };
            this.getValue = function() {
                return value;
            };
            /**
             * @param {Number} oValue
             */
            this.setValue = function(oValue) {
                value = oValue;
                update();
            };
            this.enable = function() {
                isEnabled = true;
                update();
            };
            this.disable = function() {
                isEnabled = false;
                update();
            };
            /**
             * @param {String} name
             * @param {Function} handler
             */
            this.registerEvent = function(name, handler) {
                if (!this.handlers.hasOwnProperty(name)) {
                    this.handlers[name] = [];
                }
                this.handlers[name].push(handler);
            };
            /**
             * @return {Boolean}
             */
            function _sendEvent() {
                // eslint-disable-next-line prefer-rest-params
                const name = arguments[0];
                if (this.handlers.hasOwnProperty(name)) {
                    for (let i = 0; i < this.handlers[name].length; i+=1) {
                        // eslint-disable-next-line prefer-rest-params
                        this.handlers[name][i].apply(this || window, Array.prototype.slice.call(arguments, 1));
                    }
                    return true;
                }
            };
            this.sendEvent = _sendEvent;
            parent.appendChild(canvas);
            resize();
            window.addEventListener('resize', resize);
            if (typeof ResizeObserver !== 'undefined') {
                const ro = new ResizeObserver(function() {
                    resize();
                });
                ro.observe(parent);
            }
            /**
             */
            function resize() {
                const scale = window.devicePixelRatio;
                const rect = parent.getBoundingClientRect();
                canvas.style.width = rect.width + 'px';
                canvas.style.height = rect.height + 'px';
                canvas.width = Math.round(scale * rect.right) - Math.round(scale * rect.left);
                canvas.height = Math.round(scale * rect.bottom) - Math.round(scale * rect.top);
                update();
            }
        },
    };
})();
