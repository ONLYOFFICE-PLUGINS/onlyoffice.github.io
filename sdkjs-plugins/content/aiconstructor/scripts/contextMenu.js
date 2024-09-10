class ContextMenu {
    _items = [];
    _element= null;
    _visible = false;
    _targetData = null;

    constructor(items) {
        this._items = items;
        window.contextMenuCreated.push(this);

        const me = this;
        document.addEventListener("DOMContentLoaded", function() {
            me._element = document.getElementById("context-menu-id");

            window.addEventListener("click", function(e) {
                me.getVisible() && me.hide();
            });
        });
    }

    _renderOptions() {
        const me = this;

        me._element.getElementsByClassName("context-menu-options")[0].innerHTML = '';
        this._items.forEach(function(item) {
            const node = document.createElement('div');
            node.className = 'context-menu-option';
            node.textContent = item.label;

            node.addEventListener('click', function() {
                item.handler && item.handler(me._targetData);
            });

            me._element.getElementsByClassName("context-menu-options")[0].appendChild(node);
        })
    }
    show(targetData, x, y) {
        if (!this._element)
            return;

        setTimeout(function() {
            this._targetData = targetData;
            this._element.style.left = x + "px";
            this._element.style.top = y + "px";
            this._visible = true;
            this._element.style.opacity = 0;
            this._element.style.display = "block";
            if (document.body.clientHeight < y + this._element.clientHeight) {
                // if menu is going over the bottom we should show it above the cursor
                this._element.style.top = y - this._element.clientHeight + 'px';
            }
            this._element.style.opacity = 1;

            this._renderOptions();
        }.bind(this), 5);
    }
    hide(){
        this._visible = false;
        this._element.style.display = "none";
    }
    getVisible() {
        return this._visible;
    }
}


document.addEventListener("DOMContentLoaded", function() {
    document.body.insertAdjacentHTML('beforeend',
       ' <div id="context-menu-id" class="context-menu">' +
            '<ul class="context-menu-options"></ul>' +
        '</div>'
    );
});


window.ContextMenu = ContextMenu;
window.contextMenuCreated = [];



