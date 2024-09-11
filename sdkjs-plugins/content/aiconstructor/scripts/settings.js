/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

const snippetContextMenu = new ContextMenu([
	{
		label: 'Rename',
		handler: function(item) {
			snippetItemActiveModal = item;
			renameModal.show();
		}
	},
	{
		label: 'Delete',
		handler: function(item) {
			snippetsList.deleteItem(item.guid);
		}
	},
	{
		label: 'Copy',
		handler: function(item) {
			snippetsList.copyItem(item.guid);
		}
	}
]);

const timelineContextMenu = new ContextMenu([
	{
		label: 'Select',
		handler: function(item) {
			snippetsList.selectItem(item.guid, true);
		}
	},
	{
		label: 'Delete',
		handler: function(item) {
			timelineList.deleteItem(item.slot);
		}
	},
]);

const renameModal = new Modal({
	options: {
		size: {
			width: 300,
		},
		footerStyles: 'padding: 10px',
		rerenderAfterOpen: true,
	},
	content: '<input type="text" class="form-control textSelect" id="rename-input">',
	buttons: [
		{
			text: 'OK',
			primary: true,
			autoClose: true,
			handler: handlerClickModalSubmit,
		},
		{
			text: 'Close',
			autoClose: true
		},
	],
});

renameModal.addEventListener('after:open', function () {
	const inputEl = renameModal.$el().querySelector('#rename-input');
	inputEl.value = snippetItemActiveModal.name;
	inputEl.focus();
	inputEl.select();
});


window.initCounter = 0;
const countSlotsInTimeline = 4;
let snippetItemActiveModal = null;
let codeEditor = null;
const draggedState = {
	start: {
		type: '',
		element: null
	},
	end: {
		type: '',
		element: null
	},
	data: null,

	clear: function() {
		this.start.type = '';
		this.start.element = null;
		this.end.type = '';
		this.end.element = null;
		this.data = null;
	}
};

function handlerClickModalSubmit() {
	// TODO: Add validation for input
	const inputEl = renameModal.$el().querySelector('#rename-input');
	snippetsList.changeNameItem(snippetItemActiveModal.guid, inputEl.value);
}

const snippetsList = {
	_list: [],
	_selectedItem: null,

	_renderList: function() {
		const me = this;
		const listEl = document.getElementById("snippets-list");
		const contentNewBtn = '<div class="new-snippet">Add new snippet</div>';
		let content = '';

		this._list.forEach(function(item) {
			const classes = 'snippet-item' + (me._selectedItem && item.guid === me._selectedItem.guid ? ' selected' : '');
			content += '<div class="' + classes + '" draggable="true" data-guid="' + item.guid + '">' +
				'<div class="label">' + item.name + '</div>' +
				'<div class="dots"></div>' +
				'</div>';
		});

		listEl.innerHTML = content + contentNewBtn;

		//Create event handlers
		listEl.querySelector('.new-snippet').addEventListener('click', function () {
			me.createItem();
			me.selectItem(me._list[me._list.length - 1].guid);
		});
		this._createHandlers();
	},
	_createHandlers: function() {
		const me = this;
		const elements = document.getElementsByClassName('snippet-item');
		const deleteContainerEl = document.getElementById('delete-timeline-item-container');

		deleteContainerEl.addEventListener('dragenter', function (e) {
			if(draggedState.start.type !== 'timeline') return;

			deleteContainerEl.classList.add('selected');
			draggedState.end.type = 'delete';
			draggedState.end.element = deleteContainerEl;
		});

		deleteContainerEl.addEventListener('dragleave', function (e) {
			if(draggedState.start.type !== 'timeline') return;

			setTimeout(function () {
				deleteContainerEl.classList.remove('selected');
				draggedState.end.type = '';
				draggedState.end.element = null;
			}, 1);
		});

		Array.from(elements).forEach(function(el) {
			const guid = el.getAttribute('data-guid');
			const handlerOpenMenu = function(x, y) {
				const item = me._list.find(function(item) { return item.guid === guid; });
				snippetContextMenu.show(item, x, y);
			};

			el.addEventListener('dragstart', function (e) {
				const target = e.target;
				const guid = target.getAttribute('data-guid');
				const item = me._list.find(function (item) { return item.guid == guid; });

				me.selectItem(item.guid);
				draggedState.start.type = 'snippet';
				draggedState.start.element = target;
				draggedState.data = item;

				target.classList.add('dragging');
			});
			el.addEventListener('dragend', function (e) {
				const target = e.target;
				target.classList.remove('dragging');

				if(draggedState.end.type === 'timeline') {
					const indexSlot = draggedState.end.element.getAttribute('data-slot');
					timelineList.addItem(indexSlot, draggedState.data);
				}
				draggedState.clear();
			});

			el.addEventListener('click', function(e) {
				me.selectItem(guid);
			});
			el.addEventListener('contextmenu', function(e) {
				me.selectItem(guid);
				handlerOpenMenu(e.clientX, e.clientY);
			});
			el.getElementsByClassName('dots')[0].addEventListener('click', function (e) {
				e.stopPropagation();

				const targetElement = e.target;
				const rect = targetElement.getBoundingClientRect();
				handlerOpenMenu(rect.left, rect.bottom);
			});
		});
	},

	getItems: function () {
		return this._list;
	},
	addItems: function (items) {
		const me = this;
		items.forEach(function (item) {
			me._list.push(item);
		});
		this._renderList();
		storage.save();
	},
	copyItem: function (guid) {
		const item = this._list.find(function(item) {return item.guid === guid});
		if(item) {
			const newItem = {};
			Object.assign(newItem, item);
			newItem.guid = createGuid();
			newItem.name += '_copy';
			this._list.push(newItem);
			this._renderList();
			this.selectItem(newItem.guid, true);
			storage.save();
		}
	},
	createItem: function () {
		const addedItem = {
			guid: createGuid(),
			name: 'Snippet ' + Number(this._list.length + 1),
			code: '(function()\n' +
				'{\n' +
				'})();'
		};
		this._list.push(addedItem);
		this._renderList();
		storage.save();
		return addedItem;
	},
	selectItem: function (guid, scrollTo = false) {
		const newSelectedItem = this._list.find(function(item) {
			return item.guid === guid;
		});
		if(newSelectedItem) {
			if(this._selectedItem) {
				document.querySelector(
					'.snippet-item[data-guid="' + this._selectedItem.guid + '"]'
				).classList.remove('selected');

				if(this._selectedItem.code !== codeEditor.getValue()) {
					this.changeCodeItem(this._selectedItem.guid, codeEditor.getValue());
				}
			}
			this._selectedItem = newSelectedItem;
			if(this._selectedItem) {
				document.querySelector(
					'.snippet-item[data-guid="' + this._selectedItem.guid + '"]'
				).classList.add('selected');
				codeEditor.setValue(this._selectedItem.code,  1);
				codeEditor.focus();

				if(scrollTo) {
					scrollToElementCenter(
						document.getElementById('snippets-container'),
						document.querySelector(
							'.snippet-item[data-guid="' + this._selectedItem.guid + '"]'
						)
					);
				}
			}
		}
	},
	changeNameItem(guid, name) {
		const item = this._list.find(function(item) { return item.guid === guid; });
		if(item) {
			document.querySelector(
				'.snippet-item[data-guid="' + guid + '"] .label'
			).innerHTML = name;
			item.name = name;
			storage.save();
		}
	},
	changeCodeItem(guid, code) {
		const item = this._list.find(function(item) { return item.guid === guid; });
		if(item) {
			item.code = code;
			storage.save();
		}
	},
	deleteItem: function (guid) {
		let deletedIndex = 0;
		this._list = this._list.filter(function (el, index) {
			if(el.guid == guid) {
				deletedIndex = index;
				return false;
			}
			return true;
		});
		if(this._selectedItem && this._selectedItem.guid === guid) {
			if(this._list[deletedIndex]) {
				this.selectItem(this._list[deletedIndex].guid);
			} else if(this._list[deletedIndex - 1]) {
				this.selectItem(this._list[deletedIndex - 1].guid);
			} else {
				this._selectedItem = null;
			}
		}
		this._renderList();
		storage.save();
	},
}

const timelineList = {
	_list: new Array(countSlotsInTimeline).fill(null),

	_renderList: function() {
		const listEl = document.getElementById("timeline-list");
		let content = '';

		this._list.forEach(function(item, index) {
			const classes = 'timeline-item' + (!item ? ' empty' : '');
			const name = (item ? item.name : 'Free slot ' + Number(index + 1));
			content +=
				'<div class="' + classes + '" data-slot="' + index + '" draggable="' + !!item + '">' + name + '</div>\n' +
				(index < countSlotsInTimeline - 1 ? '<img src="./resources/img/arrow.png" width="16" height="16" draggable="false"/>' : '');
		});

		listEl.innerHTML = content;


		//Register event handlers
		this._createHandlers();
	},
	_createHandlers: function() {
		const me = this;
		const elements = document.querySelectorAll('.timeline-item:not(.empty)');
		const containerEl = document.getElementById('timeline-container');

		containerEl.addEventListener('dragleave', function (e) {
			if(containerEl.contains(e.relatedTarget) || draggedState.end.type !== 'timeline') {
				return;
			}
			setTimeout(function () {
				draggedState.end.element && draggedState.end.element.classList.remove('drag-hovered');
				draggedState.end.type = '';
				draggedState.end.element = null;
			}, 1);
		});
		containerEl.addEventListener('dragover', function (e) {
			const cursorPosition = {
				x: e.clientX,
				y: e.clientY
			};

			let closestElement = getClosestElementToCursor(
				cursorPosition,
				document.getElementsByClassName('timeline-item')
			);
			if(draggedState.start.type === 'timeline' && draggedState.start.element === closestElement) {
				closestElement = null;
			}
			draggedState.end.element && draggedState.end.element.classList.remove('drag-hovered');

			if(closestElement) {
				draggedState.end.type = 'timeline';
				draggedState.end.element = closestElement;
				closestElement.classList.add('drag-hovered');
			} else {
				draggedState.end.type = '';
				draggedState.end.element = null;
			}
		});


		Array.from(elements).forEach(function(el) {
			const numSlot = el.getAttribute('data-slot');

			el.addEventListener('dragstart', function (e) {
				const target = e.target;
				const indexSlot = target.getAttribute('data-slot');

				draggedState.start.type = 'timeline';
				draggedState.start.element = target;
				draggedState.data = me._list[indexSlot];

				target.classList.add('dragging');
				document.getElementById('delete-timeline-item-container').classList.add('show');
			});
			el.addEventListener('dragend', function (e) {
				const target = e.target;
				const indexSlotStart = draggedState.start.element.getAttribute('data-slot');

				target.classList.remove('dragging');
				document.getElementById('delete-timeline-item-container').classList.remove('show');

				if(draggedState.end.type === 'timeline') {
					const indexSlotEnd = draggedState.end.element.getAttribute('data-slot');
					timelineList._swapItems(indexSlotStart, indexSlotEnd);
				} else if(draggedState.end.type === 'delete') {
					timelineList.deleteItem(indexSlotStart);
				}
				draggedState.clear();
			});

			el.addEventListener('contextmenu', function(e) {
				const item = me._list[numSlot];
				if(item) {
					const propItem = {};
					for (const key in item) {
						if (item.hasOwnProperty(key)) {
							propItem[key] = item[key];
						}
					}
					propItem.slot = +numSlot;
					timelineContextMenu.show(propItem, e.clientX, e.clientY);
				}
			});

			el.addEventListener('dblclick', function () {
				const item = me._list[el.getAttribute('data-slot')].guid;
				snippetsList.selectItem(item, true);
			});
		});
	},
	_swapItems: function (firstIndex, secondIndex) {
		const buffer = this._list[firstIndex];
		this._list[firstIndex] = this._list[secondIndex];
		this._list[secondIndex] = buffer;
		this._renderList();
	},

	addItem: function (index, item) {
		this._list[index] = item;
		this._renderList();
	},
	deleteItem: function (index) {
		this._list[index] = null;
		this._renderList();
	},
	getItems: function () {
		return this._list;
	}
};

const storage = {
	_key: 'asc_plugin_ai-constructor',

	save: function() {
		localStorage.setItem(
			this._key,
			JSON.stringify({
				snippetList: snippetsList.getItems()
			})
		);
	},
	load: function () {
		const jsonValue = localStorage.getItem(this._key) || '{}';
		const storage = JSON.parse(jsonValue);

		snippetsList.addItems(storage.snippetList || []);
	},
}

function insertCodeEditor() {
	function on_init_server(type)
	{
		if (type === (window.initCounter & type))
			return;
		window.initCounter |= type;
		if (window.initCounter === 3)
		{
			load_library("onlyoffice", "./libs/" + Asc.plugin.info.editorType + "/api.js");
		}
	}

	function load_library(name, url)
	{
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, true);
		xhr.onreadystatechange = function()
		{
			if (xhr.readyState == 4)
			{
				var EditSession = ace.require("ace/edit_session").EditSession;
				var editDoc = new EditSession(xhr.responseText, "ace/mode/javascript");
				codeEditor.ternServer.addDoc(name, editDoc);
			}
		};
		xhr.send();
	}

	function setStyles() {
		var styleTheme = document.createElement('style');
		styleTheme.type = 'text/css';

		// TODO: Refactoring. Calculate width, height and position
		window.lockTooltipsPosition = true;
		var editor_elem = document.getElementById("code-container");
		var rules = ".Ace-Tern-tooltip {\
				box-sizing: border-box;\
				max-width: " + 476 + "px !important;\
				min-width: " + 476 + "px !important;\
				box-shadow: none !important;\
				border-radius: 0px !important;\
				left: " + 323 + "px !important;\
				bottom: 80px !important;\
				}";
		// bottom: " + parseInt(document.getElementsByClassName("divHeader")[0].offsetHeight) + "px !important;\

		styleTheme.innerHTML = rules;
		document.getElementsByTagName('head')[0].appendChild(styleTheme);
	}

	codeEditor = ace.edit("code-editor");
	codeEditor.session.setMode("ace/mode/javascript");
	codeEditor.container.style.lineHeight = "20px";
	codeEditor.setValue("");

	codeEditor.getSession().setUseWrapMode(true);
	codeEditor.getSession().setWrapLimitRange(null, null);
	codeEditor.setShowPrintMargin(false);
	codeEditor.$blockScrolling = Infinity;

	ace.config.loadModule('ace/ext/tern', function () {
		codeEditor.setOptions({
			enableTern: {
				defs: ['browser', 'ecma5'],
				plugins: { doc_comment: { fullDocs: true } },
				useWorker: !!window.Worker,
				switchToDoc: function (name, start) {},
				startedCb: function () {
					on_init_server(1);
				},
			},
			enableSnippets: false
		});
	});

	if (!window.isIE) {
		ace.config.loadModule('ace/ext/language_tools', function () {
			codeEditor.setOptions({
				enableBasicAutocompletion: false,
				enableLiveAutocompletion: true
			});
		});
	}

	ace.config.loadModule('ace/ext/html_beautify', function (beautify) {
		codeEditor.setOptions({
			autoBeautify: true,
			htmlBeautify: true,
		});
		window.beautifyOptions = beautify.options;
	});

	setStyles();
	on_init_server(2);
}

window.Asc.plugin.init = function() {
	insertCodeEditor();
	storage.load();
	let firstSnippetItem = snippetsList.getItems()[0];
	if(!firstSnippetItem) {
		firstSnippetItem = snippetsList.createItem();
	}

	snippetsList.selectItem(firstSnippetItem.guid);
	timelineList._renderList();
}

window.Asc.plugin.event_onContextMenuShow = function(options) {
	console.log(options);
	switch (options.type)
	{
		case "Target":
		{
			this.executeMethod("AddContextMenuItem", [{
				guid : this.guid,
				items : [
					{
						id : "onClickItem1",
						text : { en : "Item 1", de : "Menü 1" },
						items : [
							{
								id : "onClickItem1Sub1",
								text : { en : "Subitem 1", de : "Untermenü 1" },
								disabled : true
							},
							{
								id : "onClickItem1Sub2",
								text : { en : "Subitem 2", de : "Untermenü 2" },
								separator: true
							}
						]
					},
					{
						id : "onClickItem2",
						text : { en : "Item 2", de : "Menü 2" }
					}
				]
			}]);
			break;
		}
		case "Selection":
		{
			this.executeMethod("AddContextMenuItem", [{
				guid : this.guid,
				items : [
					{
						id : "onClickItem3",
						text : { en : "Item 3", de : "Menü 3" }
					}
				]
			}]);
			break;
		}
		case 'Image':
		case 'Shape':
		{
			this.executeMethod("AddContextMenuItem", [{
				guid : this.guid,
				items : [
					{
						id : "onClickItem4",
						text : { en : "Item 4", de : "Menü 4" }
					}
				]
			}]);
			break;
		}
		default:
			break;
	}
};

window.Asc.plugin.event_onToolbarMenuClick = function(id)
{
	console.log(id);
};

document.addEventListener("DOMContentLoaded", function(e) {
	window.addEventListener("contextmenu", function(e) {
		e.preventDefault();
	});

	document.getElementById('submit-btn').addEventListener('click', function(e) {
		const result = timelineList.getItems()
			.filter(function (item) { return item })
			.map(function(item) {
				return item.code;
			});
		console.log(result);
	});
});


// Helpers
function createGuid (a,b){
	for(b=a='';a++<36;b+=a*51&52?(a^15?8^Math.random()*(a^20?16:4):4).toString(16):'');
	return b
};

function isPointInsideContainer(cursorPosition, container) {
	const containerCoord = container.getBoundingClientRect();
	return (
		cursorPosition.x >= containerCoord.left &&
		cursorPosition.x <= containerCoord.right &&
		cursorPosition.y >= containerCoord.top &&
		cursorPosition.y <= containerCoord.bottom
	);
}

function getDistanceBetweenCursorAndElement(cursorPosition, element) {
	const elementCoord = element.getBoundingClientRect();
	const elementCenter = {
		x: elementCoord.x + elementCoord.width * 0.45,
		y: elementCoord.y + elementCoord.height * 0.45
	};
	return Math.sqrt(Math.pow(cursorPosition.x - elementCenter.x, 2) + Math.pow(cursorPosition.y - elementCenter.y, 2) )
}

function getClosestElementToCursor(cursorPosition, elements) {
	let candidate = {
		el: null,
		distance: 0
	};
	Array.from(elements).forEach(function (el) {
		const distance = getDistanceBetweenCursorAndElement(cursorPosition, el);

		if(!candidate.el || distance < candidate.distance) {
			candidate.el = el;
			candidate.distance = distance;
		}
	});
	return candidate.el;
}

function scrollToElementCenter(container, element) {
	const containerHeight = container.clientHeight;
	const elementHeight = element.clientHeight;
	const elementOffsetTop = element.offsetTop;
	const scrollPosition = elementOffsetTop - (containerHeight / 2) + (elementHeight / 2);
	container.scrollTo({
		top: scrollPosition,
		behavior: 'smooth'
	});
}
