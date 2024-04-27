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
	window.Asc.plugin.init = function() {
		
		document.getElementById('btn_review').onclick = function() {
			window.Asc.plugin.sendToPlugin("onWindowMessage", {type: 'onBtnReview'});
		};
	
		document.getElementById('btn_below').onclick = function() {
			window.Asc.plugin.sendToPlugin("onWindowMessage", {type: 'onBtnBelow'});
		}
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
		let styleTheme = document.createElement('style');
		styleTheme.type = 'text/css';
		styleTheme.innerHTML = rule;
		document.getElementsByTagName('head')[0].appendChild(styleTheme);
	};

})(window, undefined);
