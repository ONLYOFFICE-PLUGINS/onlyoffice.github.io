let settingsWindow = null;

window.Asc.plugin.init = function() {
    insertTabInToolbar();
};

function insertTabInToolbar() {
    const toolbarMenuItem = {
        "id": "aiconstructor-settings",
        "type": "button",
        "text": "Settings",
        "hint": "Settings",
        "icons": "resources/light/icon.png"
    };
    const toolbarMenuTab = {
        "id": "aiconstructor",
        "text": "AI Constructor",
        "items": [toolbarMenuItem]
    };
    const toolbarMenuMainItem = {
        "guid": "asc.{89643394-0F74-4297-9E0B-541A67F1E98C}",
        "tabs": [toolbarMenuTab]
    };
    window.Asc.plugin.executeMethod ("AddToolbarMenuItem", [toolbarMenuMainItem]);
    window.Asc.plugin.attachToolbarMenuClickEvent('aiconstructor-settings', onClickSettings);
}


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
        settingsWindow.attachEvent("onWindowMessage", function(message) {
            console.log(message);
        });
    }
    settingsWindow.show(variation);
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
