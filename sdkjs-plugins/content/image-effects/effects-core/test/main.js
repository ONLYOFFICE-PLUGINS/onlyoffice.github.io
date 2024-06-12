/* eslint-disable no-invalid-this */
/* eslint-disable linebreak-style */
/* eslint-disable indent */
/* eslint-disable new-cap */
/* eslint-disable max-len */

(function(window, undefined) {
    window.Asc.plugin.onTranslate = function() {
        const labels = document.querySelectorAll('.onTranslate');
        Array.prototype.forEach.call(labels, function(item) {
            item.innerHTML = window.Asc.plugin.tr(item.innerHTML);
        });
    };
    window.ImageEffects.isWorker = true;
    let isImageLoaded = false;
    const format = {
            width: 663,
            height: 373
    };
    /**
     * @param {HTMLDivElement} effectForm
     * @param {Boolean} isCheckboxOnly
     * @param {Settings} settings
     */
    function CForm(effectForm, isCheckboxOnly, settings) {
        this.id = effectForm.id;
        this.checkbox = effectForm.firstElementChild;
        this.isCheckboxOnly = isCheckboxOnly;
        if (!this.isCheckboxOnly) {
            this.valueText = effectForm.lastElementChild;
            this.slider = new SliderModule.CSlider(settings, this.valueText.previousElementSibling);

            this.changeValue = function() {
                this.valueText.textContent = this.slider.parent.getAttribute('value');
                setEffect(false);
            };
            this.getValue = function() {
                return +this.slider.parent.getAttribute('value');
            };
            this.slider.registerEvent('change', this.changeValue.bind(this));
        }
        /**
         * @param {Event} event
         */
        function changeCheckbox(event) {
            if (!this.isCheckboxOnly) {
                if (this.checkbox.checked) {
                    this.slider.enable();
                } else {
                    this.slider.disable();
                }
            }
            setEffect(false);
        }
        this.checkbox.addEventListener('click', changeCheckbox.bind(this), false);
    }
    // eslint-disable-next-line no-unused-vars
    /**
     * @type {CForm[]}
     */
    const forms = [];
    const dropArea = document.getElementById('drop-area');
    Array.prototype.forEach.call(document.querySelectorAll('.effect-form'), (function(element) {
        if (document.querySelector('#' + element.id + '> .sliderContainer') == null) {
            forms.push(new CForm(element, true, {}));
        } else {
            forms.push(new CForm(element, false, {}));
        }
    }));

    /**
     * @type {HTMLCanvasElement}
     */
    const originCanvas = document.createElement('canvas');
    window.ImageEffects.originContext = originCanvas.getContext('2d');
    /**
     * @type {HTMLCanvasElement}
     */
    const effectCanvas = document.getElementById('previewEffect');
    window.ImageEffects.effectContext = effectCanvas.getContext('2d');
    const bg = new CBackground(effectCanvas);
    /**
     * @type {ImageData}
     */
    let middlewareImageData;
    /* Global buttons Events*/

    /* origin -> middleware -> effect */

    /**
     * @param {Event} event
     */
    function preventDefaults(event) {
        event.preventDefault();
        event.stopPropagation();
    }

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function(eventName) {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });
    /**
     * @param {Boolean} isOrigin
     */
    function setEffect(isOrigin) {
        if (window.ImageEffects.effectImageData) {
            window.ImageEffects.effectImageData.data.set(middlewareImageData.data);
            // Формируем пачку эффектов и отдаем ее
            const effects = [];
            forms.forEach(function(element) {
                if (element.isCheckboxOnly) {
                    effects.push({
                        type: element.id,
                        level: element.checkbox.checked ? 1 : 0,
                    });
                } else {
                    if (element.checkbox.checked) {
                        effects.push({
                            type: element.id,
                            level: element.getValue(),
                        });
                    }
                }
            });
            if (isOrigin) {
                window.ImageEffects.Apply({effects: effects, data: window.ImageEffects.originImageData}, true);
                if (!window.ImageEffects.isWorker) {
                    window.ImageEffects.originContext.putImageData(window.ImageEffects.originImageData, 0, 0);
                    window.ImageEffects.onExit();
                }
            } else {
                window.ImageEffects.Apply({effects: effects, data: window.ImageEffects.effectImageData}, false);
                if (!window.ImageEffects.isWorker) {
                    window.ImageEffects.effectContext.putImageData(window.ImageEffects.effectImageData, 0, 0);
                }
            }
        } else {
            throw new Error('No Image!');
        }
    }

    /* Drag n Drop */
    /**
     */
    function resize() {
        const aspect = (effectCanvas.width / effectCanvas.height);
        const rect = dropArea.parentElement.getBoundingClientRect();
        const stRect = format.height * aspect;
        const panels = document.getElementById('effects-content');
        if (panels.getBoundingClientRect().width < 510) {
            panels.style.flexDirection = 'column';
            panels.style.alignItems = 'center';
            if (panels.getBoundingClientRect().width < 303) {
                panels.style.alignItems = 'flex-start';
            }
        } else {
            panels.style.flexDirection = 'row';
            panels.style.alignItems = 'flex-start';
        }
        if (isImageLoaded) {
            if (rect.width < stRect) {
                dropArea.style.width = rect.width + 'px';
                dropArea.style.height = rect.width / aspect + 'px';
            } else {
                dropArea.style.width = format.height * aspect + 'px';
                dropArea.style.height = format.height + 'px';
            }
            effectCanvas.style.width = dropArea.style.width;
            effectCanvas.style.height = dropArea.style.height;
            return;
        }
        return;
    }
    /**
     * @param {HTMLImageElement} preImage
     */
    const imagePreview = function drawImageOnDisplay(preImage) {
        window.ImageEffects.originContext.clearRect(0, 0, originCanvas.width, originCanvas.height);
        window.ImageEffects.originContext.drawImage(preImage, 0, 0, originCanvas.width, originCanvas.height);

        window.ImageEffects.effectContext.clearRect(0, 0, effectCanvas.width, effectCanvas.height);
        window.ImageEffects.effectContext.drawImage(preImage, 0, 0, effectCanvas.width, effectCanvas.height);

        const aspect = preImage.width / preImage.height;
        dropArea.style.height = format.height + 'px';
        dropArea.style.width = format.height * aspect + 'px';

        effectCanvas.style.height = dropArea.style.height;
        effectCanvas.style.width = dropArea.style.width;
    };
    /**
     */
    function enableCheckbox() {
        const checkboxes = document.querySelectorAll('.checkbox');
        Array.prototype.forEach.call(checkboxes, (function(checkbox) {
            checkbox.removeAttribute('disabled');
        }));
    }
    const noImageText = document.getElementById('no-image-text');
    /**
     * @param {HTMLImageElement} file
     */
    function handleFiles(file) {
        noImageText.parentElement.style.display = 'none';
        effectCanvas.style.display = 'block';
        /**
         * @param {ArrayBuffer} typedArray
         * @param {String} mimeType
         * @return {String}
         */
        const url = file.src;
        const image = new Image();
        image.src = url;

        image.onload = function() {
            isImageLoaded = true;
            originCanvas.width = image.width;
            originCanvas.height = image.height;

            effectCanvas.width = image.width;
            effectCanvas.height = image.height;

            imagePreview(image);

            window.ImageEffects.originImageData = window.ImageEffects.originContext.getImageData(0, 0, originCanvas.width, originCanvas.height);
            window.ImageEffects.effectImageData = window.ImageEffects.effectContext.getImageData(0, 0, effectCanvas.width, effectCanvas.height);

            middlewareImageData = window.ImageEffects.effectContext.createImageData(window.ImageEffects.effectImageData);
            middlewareImageData.data.set(window.ImageEffects.effectImageData.data);
            enableCheckbox();
            bg.enable();
        };
    };
    window.Asc.plugin.init = function(iHtml) {
        window.ImageEffects.useOld = false;
        const wrapper = document.createElement('span');
        wrapper.innerHTML = iHtml;
        console.log(wrapper);
        if (wrapper.querySelector('img') == null || wrapper.querySelectorAll('img').length > 1) {
            this.resizeWindow(800, 413, 800, 413, 800, 413);
            return;
        }
        document.getElementById('effects-content').style.display = 'flex';
        this.resizeWindow(800, 800, 300, 700, 1920, 1080);
        window.Asc.plugin.executeMethod("GetVersion", [], function(version) {
            const ver = version.split('.');
            if (+ver[0]*10 + +ver[1] < 72) {
                window.ImageEffects.useOld = true;
                handleFiles(wrapper.querySelector('img'));
            } else {
                window.Asc.plugin.executeMethod("GetImageDataFromSelection", [], function(data) {
                    if (data === undefined) {
                        window.ImageEffects.useOld = true;
                        handleFiles(wrapper.querySelector('img'));
                        return;
                    }
                    const img = document.createElement('img');
                    img.src = data.src;
                    window.Asc.scope.width = data.width;
                    window.Asc.scope.height = data.height;
                    img.onload = function() {
                        handleFiles(img);
                    }
                });
            }
        })
    };
    window.Asc.plugin.button = function(id) {
        if (id == 0 && typeof window.ImageEffects.effectImageData !== 'undefined') {
            saveImage();
        } else {
            this.executeCommand('close', '');
        }
    };
    window.ImageEffects.onExit = function() {
        window.Asc.scope.dataURL = originCanvas.toDataURL();
        console.log(originCanvas.width, originCanvas.height);
        window.Asc.plugin.executeMethod("GetVersion", [], function(version) {
            const ver = version.split('.');
            if (+ver[0]*10 + +ver[1] < 72 || window.ImageEffects.useOld) {
                window.Asc.scope.width = (originCanvas.width * 9525);
                window.Asc.scope.height = (originCanvas.height * 9525);
                switch (window.Asc.plugin.info.editorType) {
                    case 'word': {
                        window.Asc.plugin.callCommand(function() {
                            const oDocument = Api.GetDocument();
                            const oParagraph = Api.CreateParagraph();
                            const arrResult = [];
                            arrResult.push(oParagraph);
                            const oImage = Api.CreateImage(Asc.scope.dataURL, Asc.scope.width, Asc.scope.height);
                            oParagraph.AddDrawing(oImage);
                            oDocument.InsertContent(arrResult);
                        }, true);
                        break;
                    }
                    case 'cell': {
                        window.Asc.plugin.callCommand(function() {
                            const oWorksheet = Api.GetActiveSheet();
                            oWorksheet.ReplaceCurrentImage(Asc.scope.dataURL, Asc.scope.width, Asc.scope.height);
                        }, true);
                        break;
                    }
                    case 'slide': {
                        window.Asc.plugin.callCommand(function() {
                            const oPresentation = Api.GetPresentation();
                            oPresentation.ReplaceCurrentImage(Asc.scope.dataURL, Asc.scope.width, Asc.scope.height);
                        }, true);
                    break;
                    }
                };
            } else {
                window.Asc.plugin.executeMethod("PutImageDataToSelection", [
                    {src: window.Asc.scope.dataURL,
                    width: window.Asc.scope.width,
                    height: window.Asc.scope.height}],
                    function() {
                        window.Asc.plugin.executeCommand('close', '');
                });
            }
        }); 
    };
    /**
     */
    function saveImage() {
        setEffect(true);
    }
    // for debug in browser
    // document.getElementById('effects-content').style.display = 'flex';
    // dropArea.addEventListener('drop', dropInput, false);
    // function dropInput(event) {
    //     const files = event.dataTransfer.files;
    //     const file = files.item(0);
    //     const reader = new FileReader();
    //     reader.readAsDataURL(file);
    //     reader.onloadend = function() {
    //         const img = document.createElement('img');
    //         img.src = reader.result;
    //         handleFiles(img);
    //     }
    // }
    window.addEventListener('resize', resize);
    window.ImageEffects.loadModule({enginePath: './effects-core/deploy/engine/'});
})(window, undefined);
