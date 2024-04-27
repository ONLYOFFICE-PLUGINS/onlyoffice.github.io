/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the 'License');
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an 'AS IS' BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function(window, undefined) {
	const elements = {};
	let oleData;
	let loader = null;

	window.Asc.plugin.init = function() {
		createLoader();
		sendPluginMessage({type: 'onGetOleData'});
		initElements();
		
		elements.btnCopy.onclick = function() {
			if (window.getSelection && document.createRange) { //Browser compatibility
				let range = document.createRange(); //range object
				range.selectNodeContents(elements.prompt);  //sets Range
				sel = window.getSelection();
				sel.removeAllRanges(); //remove all ranges from selection
				sel.addRange(range); //add Range to a Selection.
				document.execCommand("copy"); //copy
				sel.removeAllRanges(); //remove all ranges from selection
			}
		};
	};

	function initElements() {
		elements.btnCopy = document.getElementById('btn_copy');
		elements.prompt = document.getElementById('prompt');
		elements.preview = document.getElementById('img_preview');
		elements.mainContent = document.getElementById('main_content');
		elements.verList = document.getElementById('versions_list');
	};

	function sendPluginMessage(message) {
		window.Asc.plugin.sendToPlugin('onWindowMessage', message);
	};

	function createLoader() {
		$('#loader-container').removeClass( "hidden" );
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = showLoader($('#loader-container')[0], getTranslated('Loading...'));
	};

	function destroyLoader() {
		$('#loader-container').addClass( "hidden" )
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = undefined;
	};

	function getTranslated(text) {
		return window.Asc.plugin.tr(text);
	};

	function showinfo() {
		for (const date in oleData.data) {
			if (Object.hasOwnProperty.call(oleData.data, date)) {
				const version = oleData.data[date];
				let verContainer = document.createElement('div');
				verContainer.classList.add('div_version', 'pointer');
				verContainer.id = date;
				let verHeader = document.createElement('div');
				verHeader.classList.add('div_icon_and_date');
				let histIcon = document.createElement('div');
				histIcon.classList.add('icon_history', 'pointer');
				let label = document.createElement('label');
				label.classList.add('label_date', 'pointer');
				label.innerText = parseDate(date);
				let delIcon = document.createElement('div');
				delIcon.classList.add('icon_close', 'pointer'); 
				verHeader.appendChild(histIcon);
				verHeader.appendChild(label);
				verContainer.appendChild(verHeader);
				verContainer.appendChild(delIcon);
				
				delIcon.onclick = function(event) {
					event.stopPropagation();
					let parent = event.target.parentElement;
					if (version.active) {
						if (parent.previousElementSibling)
							parent.previousElementSibling.onclick();
						else if (parent.nextElementSibling)
							parent.nextElementSibling.onclick();
						else {
							elements.prompt.innerText = '';
							elements.preview.classList.add('hidden');
						}
					}
					delete oleData.data[date];
					parent.remove();
				};
				
				verContainer.onclick = function() {
					let oldActive = document.querySelector('.version_active');
					oldActive.classList.remove('version_active');
					this.classList.add('version_active');
					oleData.data[oldActive.id].active = false;
					version.active = true;
					elements.prompt.innerText = version.prompt;
					elements.preview.setAttribute('src', version.src)
				};
				
				elements.verList.appendChild(verContainer);
				if (version.active) {
					verContainer.classList.add('version_active');
					elements.prompt.innerText = version.prompt;
					elements.preview.setAttribute('src', version.src)
				}
				
			}
		}
		elements.mainContent.classList.remove('hidden');
		destroyLoader();
	};

	function parseDate(timestamp) {
		const date = new Date(Number(timestamp));

		const options = { 
			day: 'numeric', 
			month: 'long', 
			year: 'numeric', 
			hour: 'numeric', 
			minute: 'numeric', 
			hour12: false 
		};

		const formattedDate = date.toLocaleString('en-US', options);
		return formattedDate;
	};

	window.Asc.plugin.onTranslate = function() {
		let elements = document.querySelectorAll('.i18n');
		elements.forEach(function(element) {
			element.innerText = window.Asc.plugin.tr(element.innerText);
		})
	};

	window.Asc.plugin.onThemeChanged = function(theme) {
		// todo в окна плагина не приходит сообщение о смене темы, пока единственный вариант - посылать это событие с темой из основного плагина (но на старте оно приходит)
		window.Asc.plugin.onThemeChangedBase(theme);
		let rule = '.normal_bg { background-color: ' + theme['background-normal'] + ' !important; }';
		rule += '.wr_br { border-color: ' + theme['border-toolbar'] + ' !important; }';
		rule += '.div_version:hover { background-color: ' + theme['highlight-button-hover'] + ' !important; }';
		rule += '.version_active { background-color: ' + theme['highlight-button-pressed'] + ' !important; }';
		let styleTheme = document.createElement('style');
		styleTheme.type = 'text/css';
		styleTheme.innerHTML = rule;
		document.getElementsByTagName('head')[0].appendChild(styleTheme);
	};

	window.Asc.plugin.attachEvent('onData', function(data) {
		oleData = JSON.parse(data);
		oleData.data = JSON.parse(oleData.data);
		showinfo();
	});

	window.Asc.plugin.attachEvent('onRestore', function() {
		createLoader();
		let active = document.querySelector('.version_active');
		sendPluginMessage({type: 'onRestoreOle', data: JSON.stringify(oleData), active: (active ? active.id : null) });
	});

})(window, undefined);
