console.log("Started");

chrome.runtime.onMessage.addListener(onExternalMessage);
chrome.runtime.onMessageExternal.addListener(onExternalMessage);

function onExternalMessage (message, sender, sendResponse) {
    if (!message || !message.type)
        return;
    var type = message.type;

    console.log("Processing external message", type);

    switch (type) {
        case "reloadExtension":
            reloadExtension(message.id);
            break;
    }
};

function reloadExtension (id) {
    // FIXME: Reloading an extension this way does not update the version number and other manifest info.

    console.log("Reloading extension", id);
    chrome.management.setEnabled(id, false, function () {
        console.log("Extension disabled");
        chrome.management.setEnabled(id, true, function () {
            console.log("Extension re-enabled. Reload complete.");
        });
    });
};