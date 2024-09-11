let settingsWindow = null;
var buttonCounter = 0;

window.Asc.plugin.init = function() {
    window.Asc.plugin.executeMethod ("AddToolbarMenuItem", [toolbarMenuMainItem]);
    window.Asc.plugin.attachToolbarMenuClickEvent('aiconstructor-settings', onClickSettings);

    window.Asc.plugin.attachToolbarMenuClickEvent("button-constructor", onButtonConstructor);
};

function onClickSettings() {
    let location  = window.location;
    let start = location.pathname.lastIndexOf('/') + 1;
    let file = location.pathname.substring(start);

    // default settings for modal window (I created separate settings, because we have many unnecessary field in plugin variations)
    let variation = {
        url : location.href.replace(file, 'settings.html'),
        description : window.Asc.plugin.tr('Settings'),
        isVisual : true,
        buttons : [],
        isModal : true,
        EditorsSupport : ["word", "slide", "cell"],
        size : [800, 600]
    };

    if (!settingsWindow) {
        settingsWindow = new window.Asc.PluginWindow();
        settingsWindow.attachEvent("onAddButton", function(info) {
            buttonCounter++;
            toolbarMenuMainItem["tabs"][0]["items"].push({ 
                id: "button-constructor",
                type: "big-button",
                text: "Button" + buttonCounter,
                hint: "Button" + buttonCounter,
                data: info.data,
                lockInViewMode: true,
                enableToggle: false,
                separator: false
            });
            Asc.plugin.executeMethod("AddToolbarMenuItem", [toolbarMenuMainItem]);
        });
    }
    settingsWindow.show(variation);
}

async function onButtonConstructor(data)
{
    let arrFunctions = JSON.parse(data);

    let entryPoint = undefined;
    for (let i = 0; i < arrFunctions.length; i++)
    {
        entryPoint = await eval(arrFunctions[i]);
    }
}

window.Asc.plugin.button = function(id, windowId) {
	if (!settingsWindow)
		return;

	if (windowId) {
		switch (id) {
			case -1:
			default:
				window.Asc.plugin.executeMethod('CloseWindow', [windowId]);
		}
	}
};
