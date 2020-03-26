/*chrome.browserAction.onClicked.addListener(function() {
  //https://stackoverflow.com/questions/5345435/pop-up-window-center-screen
  var popupWidth = 520;
  var popupHeight = 720;
  var left = (screen.width / 2) - (popupWidth / 2);
  var top = (screen.height / 2) - (popupHeight / 2);

  chrome.windows.create({
    height: popupHeight,
    left: Math.round(left),
    top: Math.round(top),
    type: "popup",
    url: chrome.runtime.getURL("window.html"),
    width: popupWidth
  });
});*/

chrome.runtime.onInstalled.addListener(function() {
  chrome.declarativeContent.onPageChanged.removeRules(undefined, function() {
    chrome.declarativeContent.onPageChanged.addRules([{
      conditions: [new chrome.declarativeContent.PageStateMatcher({
        pageUrl: {schemes: ['http', 'https']},
      })
      ],
          actions: [new chrome.declarativeContent.ShowPageAction()]
    }]);
  });
});