/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *	 http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function (window, undefined) {
	let loader;
	let elements = {};
	let apiKey = '';
	let errTimeout = null;
	let tokenTimeot = null;
	let modalTimeout = null;
	let bCreateLoader = true;
	let maxTokens = 4000;
	let prefix = { previous: 'gpt-3.5', last: 'gpt-4'};

	window.Asc.plugin.init = function() {
		sendPluginMessage({type: 'onWindowReady'});
		bCreateLoader = true;
		addSlidersListeners();
		addTitlelisteners();
		initElements();

		elements.inpLenSl.oninput = onSlInput;
		elements.inpTempSl.oninput = onSlInput;
		elements.inpTopSl.oninput = onSlInput;

		elements.textArea.oninput = function(event) {
			elements.textArea.classList.remove('error_border');
			if (tokenTimeot) {
				clearTimeout(tokenTimeot);
				tokenTimeot = null;
			}
			tokenTimeot = setTimeout(function() {
				let text = event.target.value.trim();
				let tokens = window.Asc.OpenAIEncode(text);
				elements.lbTokens.innerText = tokens.length;
				checkLen();
			}, 250);

		};

		elements.textArea.onkeydown = function(e) {
			if ( (e.ctrlKey || e.metaKey) && e.key === 'Enter') {
				submitHandler();
			}
		};

		elements.divTokens.onmouseenter = function() {
			elements.modal.classList.remove('hidden');
			if (modalTimeout) {
				clearTimeout(modalTimeout);
				modalTimeout = null;
			}
		};

		elements.divTokens.onmouseleave = function() {
			modalTimeout = setTimeout(function() {
				elements.modal.classList.add('hidden');
			},100)
		};

		elements.modal.onmouseenter = function() {
			if (modalTimeout) {
				clearTimeout(modalTimeout);
				modalTimeout = null;
			}
		};

		elements.modal.onmouseleave = function() {
			elements.modal.classList.add('hidden');
		};

		elements.labelMore.onclick = function() {
			elements.linkMore.click();
		};
	};

	function submitHandler() {
		if (!apiKey)
			return;

		let settings = getSettings();
		if (settings.error || elements.textArea.classList.contains('error_border')) {
			if (Number(elements.lbAvalTokens.innerText) < 0) {
				setError('Too many tokens in your request.', false, false);
			}
			elements.textArea.classList.add('error_border');
			return;
		};
		createLoader();
		let url = `https://api.openai.com/v1${ (settings.isChat ? '/chat' : '') }/completions`;
		delete settings.isChat;

		fetch(url, {
			method: 'POST',
			headers: {
				'Content-Type': 'application/json',
				'Authorization': 'Bearer ' + apiKey,
			},
			body: JSON.stringify(settings),
		})
		.then(function(response) {
			return response.json()
		})
		.then(function(data) {
			if (data.error)
				throw data.error

			let text = data.choices[0].text || data.choices[0].message.content;
			if (!text.includes('</')) {
				// it's necessary because "PasteHtml" method ignores "\n" and we are trying to replace it on "<br>" when we don't have a html code in answer
				
				if (text.startsWith('\n'))
					text = text.replace('\n\n', '\n').replace('\n', '');
				
				text = text.replace(/\n\n/g,'\n').replace(/\n/g,'<br>');
			}

			window.Asc.plugin.executeMethod('PasteHtml', ['<div>' + text + '</div>'], function() {
				window.Asc.plugin.executeMethod('CloseWindow', [window.Asc.plugin.windowID]);
			});
		})
		.catch(function(error) {
			let bQuotaErr = error.type === 'insufficient_quota';
			setError(error.message, false, bQuotaErr);
			console.error('Error:', error);
		}).finally(function(){
			destroyLoader();
		});
	};

	function initElements() {
		elements.inpLenSl       = document.getElementById('inp_len_sl');
		elements.inpTempSl      = document.getElementById('inp_temp_sl');
		elements.inpTopSl       = document.getElementById('inp_top_sl');
		elements.inpStop        = document.getElementById('inp_stop');
		elements.textArea       = document.getElementById('textarea');
		elements.divContent     = document.getElementById('div_content');
		elements.mainError      = document.getElementById('div_err');
		elements.mainErrorLb    = document.getElementById('lb_err');
		elements.lbTokens       = document.getElementById('lb_tokens');
		elements.divTokens      = document.getElementById('div_tokens');
		elements.modal          = document.getElementById('div_modal');
		elements.lbModalLen     = document.getElementById('lb_modal_length');
		elements.labelMore      = document.getElementById('lb_more');
		elements.linkMore       = document.getElementById('link_more');
		elements.divParams      = document.getElementById('div_parametrs');
		elements.lbAvalTokens   = document.getElementById('lbl_avliable_tokens');
		elements.lbUsedTokens   = document.getElementById('lbl_used_tokens');
	};

	function addSlidersListeners() {
		const rangeInputs = document.querySelectorAll('input[type="range"]');

		function handleInputChange(e) {
			let target = e.target;
			if (e.target.type !== 'range') {
				target = document.getElementById('range');
			} 
			const min = target.min;
			const max = target.max;
			const val = target.value;
			
			console.log('handleInputChange');
			target.style.backgroundSize = (val - min) * 100 / (max - min) + '% 100%';
		};

		rangeInputs.forEach(function(input) {
			input.addEventListener('input', handleInputChange);
		});
	};

	function addTitlelisteners() {
		let divs = document.querySelectorAll('.icon_info');
		divs.forEach(function(div) {
			div.addEventListener('click', function (event) {
				let elem = event.target.parentElement.parentElement;
				elem.children[0].classList.remove('hidden');
				let top = elem.offsetTop - elem.children[0].clientHeight - 5 + 'px;';
				elem.children[0].setAttribute('style', 'top:' + top);
			});

			div.addEventListener('mouseleave', function (event) {
				event.target.parentElement.parentElement.children[0].classList.add('hidden');
			});
		});
	};

	function onSlInput(e) {
		console.log('onSlInput');
		e.target.nextElementSibling.innerText = e.target.value;
		if (e.target.id == elements.inpLenSl.id)
			elements.lbModalLen.innerText = e.target.value;
	};

	function fetchModels() {
		fetch('https://api.openai.com/v1/models', {
			method: 'GET',
			headers: {
				'Authorization': 'Bearer ' + apiKey
			}
		})
		.then(function(response) {
			return response.json();
		})
		.then(function(data) {
			if (data.error)
				throw data.error;

			let arrModels = [];
			
			for (let index = 0; index < data.data.length; index++) {
				let model = data.data[index].id;
				if (model.includes(prefix.previous) || model.includes(prefix.last))
					arrModels.push( { id: model, text: model } );
			};

			$('#sel_models').select2({
				data : arrModels
			}).on('select2:select', function(e) {
				let model = $('#sel_models').val();
				maxTokens = ( model === 'gpt-4' ) ? 8000 : ( model.includes('16k') ? 16000 : 4000 );
				checkLen();
			});

			if ($('#sel_models').find('option[value=gpt-4]').length) {
				$('#sel_models').val('gpt-4').trigger('change');
			}
			elements.divContent.classList.remove('hidden');
		})
		.catch(function(error) {
			setError('Problem with models loading.', true, false);
			apiKey = '';
			console.error(error);
		}).finally(function() {
			destroyLoader();
		});
	};

	function getSettings() {
		let model = $('#sel_models').val();
		let value = elements.textArea.value.trim();
		let obj = {
			model : model,
			isChat: (model.includes('-3.5') && !model.includes('-3.5-turbo-instruct')) || model.includes('-4')
		};
		if (!value.length) {
			obj.error = true;
		} else {
			if (obj.isChat)
				obj.messages = [{role: 'user', content: value}];
			else
				obj.prompt = value;

			let temp = Number(elements.inpTempSl.value);
			obj.temperature = ( temp < 0 ? 0 : ( temp > 1 ? 1 : temp ) );
			let len = Number(elements.inpLenSl.value);
			obj.max_tokens = ( len < 0 ? 0 : ( len > maxTokens ? maxTokens : len ) );
			let topP = Number(elements.inpTopSl.value);
			obj.top_p = ( topP < 0 ? 0 : ( topP > 1 ? 1 : topP ) );
			let stop = elements.inpStop.value;
			if (stop.length)
				obj.stop = stop;
		}
		return obj;
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

	function setError(error, bConst, bQuotaErr) {
		if (bQuotaErr) {
			elements.mainErrorLb.innerHTML = '<label style="martgin-bottom:5px">' + getTranslated('You exceeded your current quota, please check your plan and billing details.') + '</label>' +
			'<p></p> <label>' + getTranslated('For more information on this error, read the docs:') + '</label>' +
			'<br> <a target="_blank" class="text-link" href="https://platform.openai.com/docs/guides/error-codes/api-errors.">https://platform.openai.com/docs/guides/error-codes/api-errors.</a>';
		} else {
			elements.mainErrorLb.innerHTML = '<label>' + getTranslated(error) + '</label>';
		}
		elements.mainError.classList.remove('hidden');
		if (errTimeout) {
			clearTimeout(errTimeout);
			errTimeout = null;
		}
		if (!bConst)
			errTimeout = setTimeout(clearMainError, 5000);
	};

	function clearMainError() {
		elements.mainError.classList.add('hidden');
		elements.mainErrorLb.innerHTML = '';
	};

	function getTranslated(key) {
		return window.Asc.plugin.tr(key);
	};

	function checkLen() {
		let cur = Number(elements.lbTokens.innerText);
		let maxLen = Number(elements.inpLenSl.value);
		let newValue = maxTokens - cur;

		if (cur > maxTokens) {
			elements.textArea.classList.add('error_border');
		} else {
			elements.textArea.classList.remove('error_border');
		}
	
		if (cur + maxLen > maxTokens) {
			setTokensLenght(newValue, newValue);
		} else {
			setTokensLenght(maxLen, newValue);
		}
	};

	function setTokensLenght(val, max) {
		elements.inpLenSl.setAttribute('max', max);
		elements.inpLenSl.value = val;
		let event = document.createEvent('Event');
		event.initEvent('input', true, true);
		elements.inpLenSl.dispatchEvent(event);
		elements.lbAvalTokens.innerText = elements.inpLenSl.getAttribute('max');
		elements.lbUsedTokens.innerText = elements.lbTokens.innerText;
	};

	function sendPluginMessage(message) {
		window.Asc.plugin.sendToPlugin("onWindowMessage", message);
	};

	window.Asc.plugin.onTranslate = function() {
		if (bCreateLoader)
			createLoader();

		let elements = document.querySelectorAll('.i18n');
		bCreateLoader = true;

		for (let index = 0; index < elements.length; index++) {
			let element = elements[index];
			element.innerText = getTranslated(element.innerText);
		}
	};

	window.Asc.plugin.onThemeChanged = function(theme) {
		// todo в окна плагина не приходит сообщение о смене темы, пока единственный вариант - посылать это событие с темой из основного плагина (но на старте оно приходит)
		window.Asc.plugin.onThemeChangedBase(theme);

		let rule = ".select2-container--default.select2-container--open .select2-selection__arrow b { border-color : " + window.Asc.plugin.theme["text-normal"] + " !important; }";
		let sliderBG, thumbBG
		if (theme.type.indexOf('dark') !== -1) {
			sliderBG = theme.Border || '#757575';
			// for dark '#757575';
			// for contrast dark #616161
			thumbBG = '#fcfcfc';
		} else {
			sliderBG = '#efefef';
			thumbBG = '#c0c0c0';
		}
		rule += '.text-link, .text-link:hover, .text-link:active, .text-link:visited{color: ' + theme['text-normal'] + ' !important;}';
		rule += '\n .normal_bg { background-color: ' + theme['background-normal'] + ' !important; }';
		rule += '\n input[type="range"] { background-color: '+sliderBG+' !important; background-image: linear-gradient('+thumbBG+', '+thumbBG+') !important; }';
		rule += '\n input[type="range"]::-webkit-slider-thumb { background: '+thumbBG+' !important; }';
		rule += '\n input[type="range"]::-moz-range-thumb { background: '+thumbBG+' !important; }';
		rule += '\n input[type="range"]::-ms-thumb { background: '+thumbBG+' !important; }';
		
		let styleTheme = document.createElement('style');
		styleTheme.type = 'text/css';
		styleTheme.innerHTML = rule;
		document.getElementsByTagName('head')[0].appendChild(styleTheme);
	};

	window.Asc.plugin.attachEvent("onApiKey", function(settings) {
		apiKey = JSON.parse(settings).apiKey;
		if (apiKey) {
			fetchModels();
		} else {
			bCreateLoader = false;
			destroyLoader();
		}
	});

	window.Asc.plugin.attachEvent("onClearBtn", function() {
		elements.textArea.value = '';
		elements.lbTokens.innerText = 0;
		elements.textArea.focus();
		checkLen();
	});

	window.Asc.plugin.attachEvent("onSubmitBtn", submitHandler);

})(window, undefined);
