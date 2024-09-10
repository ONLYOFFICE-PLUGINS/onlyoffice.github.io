
class Modal extends EventTarget {
    _elementWrapper = null;
    _element = null;
    _elementContent = null;
    _elementFooter = null;

    _visible = false;
    _options = {
        size: {
            width: 300,
            height: 'auto',
        },
        contentStyles: '',
        footerStyles: '',
        rerenderAfterOpen: false
    };
    _content = '';
    _buttons = [];

    id='';

    /**
     * @typedef ButtonObject
     * @type {Object}
     * @property {string} text - Label.
     * @property {boolean} primary - Primary style.
     * @property {boolean} autoClose - Close after click.
     * @property {string} handler - Click handler.
     */

     /**
      * @param {Object} props - Properties
      * @param {Object} props.options - Parameter object.
      * @param {Object} props.options.size - Size.
      * @param {number} props.options.size.width - Width.
      * @param {number} props.options.size.height - Height.
      * @param {string} props.options.contentStyles - Main block styles.
      * @param {string} props.options.footerStyles - Footer styles.
      * @param {boolean} props.options.rerenderAfterOpen - Need to be rerender after opening

      * @param {String} props.content - HTML content.

      * @param {Array.<ButtonObject>} props.buttons - Buttons in footer.
     */
    constructor(props) {
        super();
        this.id = this._createGuid();
        this._deepFill(this._options, props.options);
        this._content = props.content;
        this._buttons = props.buttons;
        document.addEventListener("DOMContentLoaded", function() {
            this._elementWrapper = document.createElement("div");
            this._elementWrapper.id = 'modal-' + this.id;
            this._elementWrapper.className = "modal-container";
            this._elementWrapper.innerHTML = '<div class="modal-mask"></div>';

            this._element = document.createElement("div");
            this._element.className = 'modal';
            this._elementWrapper.appendChild(this._element);

            document.body.appendChild(this._elementWrapper);

            this._element.style.width = this._options.size.width + 'px';
            this._element.style.height = (
                typeof this._options.size.height === 'number' ? this._options.size.height + 'px' : 'auto'
            );
        }.bind(this));

         document.addEventListener('keydown', function(e) {
             if (e.key == "Escape" && this._visible) {
                 this.hide();
             }
         }.bind(this));
    }

    _createGuid (a,b){
        for(b=a='';a++<36;b+=a*51&52?(a^15?8^Math.random()*(a^20?16:4):4).toString(16):'');
        return b
    }
    _renderContent() {
        if(!this._elementContent) {
            const node = document.createElement('div');
            node.className = 'modal-content';
            node.style = this._options.contentStyles;
            this._elementContent = node;
            this._element.appendChild(node);
        } else {
            this._elementContent.innerHTML = '';
        }
        this._elementContent.insertAdjacentHTML('beforeend', this._content);
        setTimeout(function() {
            this.dispatchEvent(new CustomEvent('after:open'));
        }.bind(this), 0);
    }
    _renderFooter() {
        if(this._buttons.length === 0) {
            if(this._elementFooter) {
                this._elementFooter.remove();
                this._elementFooter = null;
            }
            return;
        }

        if(!this._elementFooter) {
            const node = document.createElement('div');
            node.className = 'modal-footer';
            node.style = this._options.footerStyles;
            this._elementFooter = node;
            this._element.appendChild(node);
        } else {
            this._elementFooter.innerHTML = '';
        }

        this._buttons.forEach(function (button) {
            const node = document.createElement('button');
            node.className = 'btn-text-default';
            if(button.primary) {
                node.classList.add('primary');
            }
            node.textContent = button.text;

            node.addEventListener('click', function() {
                button.autoClose && this.hide();
                button.handler && button.handler();
            }.bind(this));
            this._elementFooter.appendChild(node);
        }.bind(this))
    }
    _deepFill(target, source) {
        for (let key in target) {
            if (target.hasOwnProperty(key)) {
                if (typeof target[key] === 'object' && target[key] !== null && !Array.isArray(target[key])) {
                    // Если свойство - объект, вызываем рекурсивно для вложенных объектов
                    if (source[key] && typeof source[key] === 'object') {
                        this._deepFill(target[key], source[key]);
                    }
                } else {
                    // Если свойство есть во втором объекте, обновляем значение
                    if (source.hasOwnProperty(key)) {
                        target[key] = source[key];
                    }
                }
            }
        }
        return target;
    }

    $el() {
        return this._element;
    }

    show() {
        if(this._options.rerenderAfterOpen || !this._elementContent) {
            this._renderContent();
            this._renderFooter();
        }

        this._elementWrapper.style.display = "flex";
        setTimeout(function () {
            this._elementWrapper.style.opacity = 1;
            this._element.style.scale = 1;
        }.bind(this), 10);
        this._visible = true;
    }
    hide() {
        this._elementWrapper.style.opacity = 0;
        this._element.style.scale = 0.5;
        this._visible = false;
        setTimeout(function () {
            this._elementWrapper.style.display = "none";
        }.bind(this), 100);
    }
    getVisible() {
        return this._visible;
    }
}

window.Modal = Modal;
