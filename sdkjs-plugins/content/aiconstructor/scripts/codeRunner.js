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

var AI = {
	models : {
		"gpt-3.5" : {
			name : "gpt-3.5-turbo",
			url : "https://api.openai.com/v1/",
			max_tokens : 3999,
			key : "YOUR KEY"
		}
	}
};

AI.callMethod = async function(name, args)
{
	return new Promise(resolve => (function(){
		Asc.plugin.executeMethod(name, args || [], function(returnValue){
			resolve(returnValue);
		});
	})());
};

AI.callCommand = async function(func)
{
	return new Promise(resolve => (function(){
		Asc.plugin.callCommand(func, false, true, function(returnValue){
			resolve(returnValue);
		});
	})());
};

function Provider(name)
{
	this.name = name;

	this.getModel = function() {
		for (let i in AI.models) {
			if (i === this.name)
				return AI.models[i];
		}
		return null;
	}

	this.request = async function(entryPoint) {
		await AI.callMethod("StartAction", ["Block", "AI..."]);

		let model = this.getModel(this.name);
		if (!model)
			return undefined;

		let headers = {};
		headers["Content-Type"] = "application/json";
		if (model.key)
			headers["Authorization"] = "Bearer " + model.key;

		let tokens_content = window.Asc.OpenAIEncode(entryPoint);

		let result = await fetch(model.url + "chat/completions", {
			method: 'POST',
			headers: headers,
			body: JSON.stringify({
				max_tokens : model.max_tokens - tokens_content.length,
				model : model.name,
				messages:[{role:"user",content:entryPoint}]
			})
		});

		let resultJson = await result.json();
		
		let text = resultJson.choices[0].message.content;
		let i = 0; let trimStartCh = "\n".charCodeAt(0);
		while (text.charCodeAt(i) === trimStartCh)
			i++;
		if (i > 0)
			text = text.substring(i);

		await AI.callMethod("EndAction", ["Block", "AI..."]);
		
		return text;
	};
}


AI.provider = function(name)
{
	return new Provider(name);
};
