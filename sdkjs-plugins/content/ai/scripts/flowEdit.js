var flowDesk = null;

var $groupsList = $('#left-panel-groups-list');


var inputsGroup = createGroupView('inputs-group', {
	type: 'input',
	title: 'Inputs',
	icon: 'btn-update'
});

var actionsGroup = createGroupView('actions-group', {
	type: '',
	title: 'Action',
	icon: 'ai-texts'
});

var outputsGroup = createGroupView('outputs-group', {
	type: 'output',
	title: 'Outputs',
	icon: 'ai-texts'
});


inputsGroup.set([
	{ name: 'One word' },
	{ name: 'Selected text' },
	{ name: 'The whole text' },
	{ name: 'Image' }
]);

actionsGroup.set([
	{ name: 'Summarize' }
]);

outputsGroup.set([
	{ name: 'Comment' },
	{ name: 'Review' },
	{ name: 'Replace' }
]);



var scrollbarList = new PerfectScrollbar("#left-panel-groups-list", {});

window.Asc.plugin.init = function() {
	window.Asc.plugin.sendToPlugin("onInit");
	window.Asc.plugin.attachEvent("onThemeChanged", onThemeChanged);
	
	insertFlowdesk();
}
window.Asc.plugin.onThemeChanged = onThemeChanged;

window.Asc.plugin.onTranslate = function () {
	let elements = document.querySelectorAll('.i18n');
	elements.forEach(function(element) {
		element.innerText = window.Asc.plugin.tr(element.innerText);
	});
};

window.addEventListener("resize", onResize);
onResize();

function insertFlowdesk() {
	let config = {
		boxId: 'flowdesk',
		miniMap: true,
		zoomControls: true,
		// uitheme: 'ui-theme-dark',
		events: {
			onLoad: e => {
			console.log('frame loaded')
			},
			onGetRoute: data => {
			console.log('on getRow answer', data)
			},
		},
	};
	flowDesk = new FlowDesk(config);
}

function onThemeChanged(theme) {
	window.Asc.plugin.onThemeChangedBase(theme);
	themeType = theme.type || 'light';
	addCssVariables(theme);
	
	var classes = document.body.className.split(' ');
	classes.forEach(function(className) {
		if (className.indexOf('theme-') != -1) {
			document.body.classList.remove(className);
		}
	});
	document.body.classList.add(theme.name);
	document.body.classList.add('theme-type-' + themeType);

	let btnIcons = document.getElementsByClassName('icon');
	for (let i = 0; i < btnIcons.length; i++) {
		let icon = btnIcons[i];
		let src = icon.getAttribute('src');
		let newSrc = src.replace(/(icons\/)([^\/]+)(\/)/, '$1' + themeType + '$3');
		icon.setAttribute('src', newSrc);
	}
}

function getZoomSuffixForImage() {
	var ratio = Math.round(window.devicePixelRatio / 0.25) * 0.25;
	ratio = Math.max(ratio, 1);
	ratio = Math.min(ratio, 2);
	if(ratio == 1) return ''
	else {
		return '@' + ratio + 'x';
	}
}

function onResize () {
	$('img').each(function() {
		var el = $(this);
		var src = $(el).attr('src');
		if(!src.includes('resources/icons/')) return;

		var srcParts = src.split('/');
		var fileNameWithRatio = srcParts.pop();
		var clearFileName = fileNameWithRatio.replace(/@\d+(\.\d+)?x/, '');
		var newFileName = clearFileName;
		newFileName = clearFileName.replace(/(\.[^/.]+)$/, getZoomSuffixForImage() + '$1');
		srcParts.push(newFileName);
		el.attr('src', srcParts.join('/'));
	});
}

function addCssVariables(theme) {
	let colorRegex = /^(#([0-9a-f]{3}){1,2}|rgba?\([^\)]+\)|hsl\([^\)]+\))$/i;

	let oldStyle = document.getElementById('theme-variables');
	if (oldStyle) {
		oldStyle.remove();
	}

	let style = document.createElement('style');
	style.id = 'theme-variables';
	let cssVariables = ":root {\n";

	for (let key in theme) {
		let value = theme[key];

		if (colorRegex.test(value)) {
			let cssKey = '--' + key.replace(/([A-Z])/g, "-$1").toLowerCase();
			cssVariables += ' ' + cssKey + ': ' + value + ';\n';
		}
	}

	cssVariables += "}";

	style.textContent = cssVariables;
	document.head.appendChild(style);
}


function createGroupView(id, options) {
	let $el = $('<div id="' + id+ '"></div>');
	let instance = new GroupView($el, options);
	$groupsList.append($el);
	return instance;
}

function GroupView($el, options) {
    this._init = function() {
        var defaults = {
			type: '',
			title: '',
			icon: ''
        };
        this.options = Object.assign({}, defaults, options);

        this.$el = $el;
		this.list = [];
		this.collapsed = false;

        this._render();
    };

	this.set = function(list) {
        this.list = list;
        this._render();
    };

	this._toggleCollapse = function() {
		this.collapsed = !this.collapsed;
		this.$el.toggleClass('collapsed', this.collapsed);

		let me = this;
		let $componentsList = $el.find('.components-list');

		$componentsList.css('height', $componentsList[0].scrollHeight + 'px');
		if(me.collapsed) {
			setTimeout(function() {
				$componentsList.css('height', '0px');
			}, 0);
		}

		this.tooltip.setText(this.collapsed ? 'Uncollapse' : 'Collapse');
	};

	this._render = function() {
		let me = this;
		this.$el.empty();
		this.$el.addClass('group');
		this.$el.append(
			'<div class="header label-row">' +
				'<div class="align-center">' + 
					'<img class="label-icon icon" src="resources/icons/light/' + this.options.icon + '.png" class="icon"/>' +
					'<label class="i18n">' + this.options.title + '</label>' +
				'</div>' +
				'<button class="btn-text-default collapse-btn not-border-btn">' + 
					'<img src="resources/icons/light/btn-demote.png" class="icon"/>' +
				'</button>' +
			'</div>'
		);
		this.$el.find('.collapse-btn').on('click', this._toggleCollapse.bind(this));

		let $componentsList = $('<div class="components-list"></div>'); 
		this.list.forEach(function(item, index) {
			let $item = $('<div class="item" draggable="true">' + item.name + '</div>');
			$item.on('dragstart', function(e) {
				let text = item.name;
				let strcode = '(async () => {return ' + text + ';})(entryPoint);';
				let obj = {
					title: text,
					code: strcode,
					type: me.options.type
				};
				
				e.originalEvent.dataTransfer.setData("info", JSON.stringify(obj));
			})
			$componentsList.append($item);
		});
		this.$el.append($componentsList);

		this.tooltip = new Tooltip(this.$el.find('.collapse-btn')[0], {
			text: 'Collapse',
			yAnchor: 'top',
			xAnchor: 'center',
			align: 'center'
		});

		// TODO: Сделать чтобы перерендеривался только список а не весь контент группы
	};

    this._init();
}
