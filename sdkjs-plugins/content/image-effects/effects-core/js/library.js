/* eslint-disable linebreak-style */
/* eslint-disable max-len */
/* eslint-disable indent */
/* eslint-disable space-before-blocks */
'use strict';

/**
 * @typedef {Object} Effect
 * @property {String} type
 * @property {Number} level
 */

/**
 * @typedef {Object} Message
 * @property {Array<Effect>} effects
 * @property {ImageData} data
 */

(function(window){
    window.ImageEffects = window.ImageEffects || {};
    window.ImageEffects.isReady = false;
    window.ImageEffects.Apply = function() {
        console.log('Module is not ready');
    };

    window.ImageEffects.onLoadModule = function(exports) {
        window.ImageEffects.isReady = true;
        window.ImageEffects.Apply = exports.ApplyEffect;
    };
    // Точка входа из main.js
    window.ImageEffects.loadModule = function(settings) {
        let url = settings.enginePath ? settings.enginePath : './effects-core/deploy/engine/';

        let useWasm = false;
        const webAsmObj = window['WebAssembly'];
        if (typeof webAsmObj === 'object') {
            if (typeof webAsmObj['Memory'] === 'function') {
                if ((typeof webAsmObj['instantiateStreaming'] === 'function') || (typeof webAsmObj['instantiate'] === 'function')) {
                    useWasm = true;
                }
            }
        }
        url += (useWasm ? 'effects.js' : 'effects_ie.js');
        // eslint-disable-next-line no-var
        if (window.ImageEffects.isWorker){
            const worker = new Worker(url);
            let isSent = false;
            let backup = null;
            /**
             * @param {Message} message
             */
            function ApplyEffect(message){
                if (isSent) {
                    backup = {
                        effects: message.effects,
                        data: window.ImageEffects.effectContext.createImageData(message.data),
                    };
                    backup.data.data.set(message.data.data);
                    return;
                }
                isSent = true;
                worker.postMessage(message);
            }
            worker.onmessage = function(e) {
                if (e.data == 'module is ready'){
                    console.log('WebWorkers will be used');
                    ImageEffects.onLoadModule({ApplyEffect: ApplyEffect});
                    console.log((useWasm ? 'wasm' : 'asmjs') + ' module will be used');
                    worker.onmessage = function(e){
                        window.ImageEffects.effectImageData.data.set(e.data.data);
                        window.ImageEffects.effectContext.putImageData(window.ImageEffects.effectImageData, 0, 0);
                        isSent = false;
                        if (backup) {
                            isSent = true;
                            this.postMessage(backup);
                            backup = null;
                        }
                    };
                } else {
                    throw new Error('Unknown message from worker:' + url);
                }
            };
        } else {
            const script = document.createElement('script');
            script.type = 'text/javascript';
            script.src = url;
            script.onload = function() {
                console.log('default script will be used');
                console.log((useWasm ? 'wasm' : 'asmjs') + ' module will be used');
            };
            script.onerror = function() {
                // TODO: попробовать загрузить еще сколько-то раз (максимальное число попыток  - зашито в коде - например 5)
            };
            document.head.appendChild(script);
        }
    };
})(self);
