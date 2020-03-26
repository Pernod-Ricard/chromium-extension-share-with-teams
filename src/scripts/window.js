//#region Azure Active Directory Client ID
const aad_clientId = '';
//#endregion

//#region Controls Instances & Constants
var isImportantControl = null;
var mentionChannelControl = null;
var mentionTeamControl = null;
var tagsBoxControl = null;
var tagsLoadIndicatorControl = null;
var teamsLoadPanelControl = null;
var teamsTreeControl = null;
var toolbarControl = null;
var txtSubjectControl = null;
var txtMessageControl = null;

const refreshButtonIndex = 0;
const shareButtonIndex = 999;
//#endregion

//#region Global Variables & Constants
var currentTab = null;
var requestSettings = null;
var selectedChannel = null;
var teamsStore = null;
var metaTitle = null;
var metaDescription = null;
var metaImage = null;
var searchEmployeesTimeout = null;
var shareButtonEnabled = true;
var teamTreeComponent = null;
var teamElementToScroll = null;

const loadingChannelsMessage = 'Loading channels...';
const loadingTeamsMessage = 'Loading your teams...';
const tabMetaScript = `(function getTabMeta(){
    const title = document.querySelector('meta[property="og:title"]') 
        ? document.querySelector('meta[property="og:title"]').content
        : undefined;
    const description = document.querySelector('meta[property="og:description"]') 
        ? document.querySelector('meta[property="og:description"]').content
        : undefined;
    var image = document.querySelector('meta[property="og:image"]')?.content;
    if(!image){
        image = document.querySelector('link[id="favicon"]')?.href;
    }
    return { title, description, image };
})()`; 
//#endregion


//#region Utilities

function changeButtonStatus(index, disabled){
    if(index == shareButtonIndex){
        shareButtonEnabled = !disabled;
    }
    else{
        var buttons = toolbarControl.option("items");
        buttons[index].options.disabled = disabled;
        toolbarControl.option('items', buttons);
    }
}

function close(){
    chrome.tabs.getCurrent(function (tab){
        if(tab){
            chrome.tabs.remove(tab.id);
        }
    });
}

function compareChannels(channelA, channelB) {
    if(!channelA){
        return -1;
    }
    if(channelA.displayName.toUpperCase() == 'GENERAL'){
        return -1;
    }
    if(channelB.displayName.toUpperCase() == 'GENERAL'){
        return 1;
    }
    return channelA.displayName.localeCompare(channelB.displayName);
}

function compareTeams(teamA, teamB) {
    if(!teamA){
        return -1;
    }
    return teamA.displayName.localeCompare(teamB.displayName);
}

function createGuid(){
    return (S4() + S4() + S4() + '4' + S4().substr(0,3) + S4() + S4() + S4() + S4()).toLowerCase();
}

function parceResponseUrlForAccessToken(responseUrl){
    const paramName = '#access_token';
    const paramIndex = responseUrl.indexOf(paramName);
    
    if (paramIndex > -1){
        var token = responseUrl.substring(paramIndex + paramName.length + 1);
        token = token.substring(0, token.indexOf('&'));
        return token;
    }
}

function S4() {
    return (((1+Math.random())*0x10000)|0).toString(16).substring(1); 
}

function truncateText(text){
    if(!text){
        return text;
    }
    if(text.length <= 32){
        return text;
    }
    return text.substring(0, 32) + '...';
}

//#endregion

//#region Cache

function cacheHasExpired(cacheDate){
    console.log('Cache date : ' + cacheDate);
    var age = Math.round((cacheDate - new Date()) / (1000*60*60*24));
    console.log('Cache date : ' + age + ' day(s)');
    var hasExpired = age > 10;
    console.log('Cache has expired : ' + hasExpired);
    return hasExpired;
}

function clearCache(){
    chrome.storage.local.clear();
    console.log('The cache have been cleared');
}

function saveInCache(){
    console.log('Save teams in cache');
    clearCache();
    teamsStore.load()
        .done(function (data) {
            chrome.storage.local.set({ date: (new Date()).toJSON() });
            chrome.storage.local.set({ teams: data });
            console.log('Teams have been saved (for today)');
        })
        .fail(function (error) {
            console.error(error);
        });
}

//#endregion

//#region Teams

function createMessage(){
    const cardGuid = createGuid();
    const cardContent = {
        title: metaTitle ?? currentTab.title,
        subtitle:  (new URL(currentTab.url)).hostname,
        text: metaDescription, 
        images: [
            { 
                "url": metaImage,
                "alt": metaTitle ?? currentTab.title
            }
        ],
        tap: {
            title: metaTitle ?? currentTab.title,
            type: "openUrl",
            value: currentTab.url
        }
    };

    var importance = 'normal'
    if(isImportantControl.option("value")){
        importance = 'high'
    }
    var mentionTeam = mentionTeamControl.option('value');
    var mentionChannel = mentionChannelControl.option('value');
    var mentionnedPeople = tagsBoxControl.option('value');
    var additionalMessage = txtMessageControl.option('value');
    var mentionsCounter = 0;
    var mentions = new Array();

    var message = '<div style="width:100%;">';
    if(mentionTeam || mentionChannel){
        message += '<div>';
        if(mentionTeam){
            message += '<at id="' + mentionsCounter + '">' + teamsTreeControl.selectedItem.teamName + '</at>';
            mentions.push(
                { 
                    id: mentionsCounter, 
                    mentionText: teamsTreeControl.selectedItem.teamName, 
                    mentioned: 
                    { 
                        conversation:
                        { 
                            id: teamsTreeControl.selectedItem.teamId, 
                            displayName: teamsTreeControl.selectedItem.teamName,
                            conversationIdentityType: 'team'
                        }
                    }
                });
            mentionsCounter++;
        }
        if(mentionChannel){
            if(mentionTeam){
                message += '&nbsp;&nbsp;';
            }
            message += '<at id="' + mentionsCounter + '">' + teamsTreeControl.selectedItem.name + '</at>';
            mentions.push(
                { 
                    id: mentionsCounter, 
                    mentionText: teamsTreeControl.selectedItem.name, 
                    mentioned: 
                    { 
                        conversation:
                        { 
                            id: teamsTreeControl.selectedItem.id, 
                            displayName: teamsTreeControl.selectedItem.name,
                            conversationIdentityType: 'channel'
                        }
                    }
                });
            mentionsCounter++;
        }
        message += '</div>';
    }
    if(mentionnedPeople){
        message += '<div>';
        for (const mentionned of mentionnedPeople){
            message += '<at id="' + mentionsCounter + '">' + mentionned.displayName + '</at>';
            mentions.push(
                { 
                    id: mentionsCounter, 
                    mentionText: mentionned.displayName, 
                    mentioned: 
                    { 
                        user:
                        { 
                            id: mentionned.id, 
                            displayName: mentionned.displayName,
                            userIdentityType: 'aadUser'
                        }
                    }
                });
                mentionsCounter++;
        }
        message += '</div>';
    }
    if(additionalMessage){
        message += '<div>'
                 + additionalMessage
                 + '</div>';
    }

    message += '<div>'
              +'<a href="' + currentTab.url + '" title="' + currentTab.url + '" target="_blank" rel="noreferrer noopener">' + currentTab.url + '</a>'
              +'</div>'
              +'<div style="width:100%;">'
              +'<attachment id="' + cardGuid + '"></attachment>'
              +'</div>'
              +'</div>';

    return {
        body: {
            content: message,
            contentType: 'html'
        },
        attachments:[
            {
                id: cardGuid,
                contentType: 'application/vnd.microsoft.card.thumbnail',
                content: JSON.stringify(cardContent)
            }
        ],
        importance: importance,
        mentions: mentions,
        subject: txtSubjectControl.option('value')
    };
}


function getTeams(ignoreCache){
    chrome.identity.launchWebAuthFlow(
    {
        url: 'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?' 
           + 'client_id=' + aad_clientId
           + '&scope=user.read.all group.readwrite.all'
           + '&response_type=token',
        interactive: true
    }, 
    function (responseUrl){
        if(!responseUrl){
            console.warn('Authentication has failed or has been canceled by user');
            DevExpress.ui.notify('Authentication has failed or has been canceled by user', 'warning', 1500);
            teamsLoadPanelControl.hide();
            return;
        }
        const graphAccessToken = parceResponseUrlForAccessToken(responseUrl);
        const reqHeaders = new Headers();
        reqHeaders.append('Authorization', 'Bearer ' + graphAccessToken);
        reqHeaders.append('Content-Type', 'application/json');
        requestSettings = { method: 'GET', headers: reqHeaders };

        if(!ignoreCache){
            loadTeamsFromCacheAsync();
        }
        else{
            console.log('Ignoring cache has been requested, loading teams from server');
            loadTeamsFromApiAsync();
        }
    });
}

async function loadTeamsFromApiAsync(){
    console.log('Loading teams from server');
    clearCache();

    if(requestSettings){
        teamsLoadPanelControl.option('message', loadingTeamsMessage);
        teamsLoadPanelControl.show();
        changeButtonStatus(refreshButtonIndex, true);
        changeButtonStatus(shareButtonIndex, true);

        const teamsRequest = new Request('https://graph.microsoft.com/v1.0/me/joinedTeams', requestSettings);
        const teamsResponse = await fetch(teamsRequest, requestSettings)
        .catch(function(error) { console.error(error.message); });

        if(teamsResponse.ok){
            //deserialize returned data
            const teamsData = await teamsResponse.json();
            teamsData.value.sort(compareTeams);

            //prepare data
            const teams = new Array();
            teamsData.value.forEach(team => {
                teams.push({ id: team.id, name: team.displayName, icon: 'icons/teams_32x32.png', expanded: false, type: 'team', items: [] });
            });

            //init the TreeView data store
            teamsStore = new DevExpress.data.ArrayStore({ key: 'id', data: teams });
            const teamsSource = new DevExpress.data.DataSource({ store: teamsStore, pushAggregationTimeout: 2000, reshapeOnPush: false });
        
            //load the teams in the TreeView
            teamsTreeControl.option('dataSource', teamsSource);
            
            //load teams icons from Graph
            await loadTeamsIconsAsync(teams);

            //load teams channels from Graph
            await loadTeamsChannelsAsync(teams);
            
            //save data in cache
            saveInCache();

            teamsLoadPanelControl.hide();
            changeButtonStatus(refreshButtonIndex, false);
            changeButtonStatus(shareButtonIndex, (teamsTreeControl.selectedItem));
        }
        else{
            console.error(teamsResponse.status);
            DevExpress.ui.notify('An error has occured. Please close then open me again', 'error', 1500);
            teamsLoadPanelControl.hide();
            changeButtonStatus(refreshButtonIndex, false);
        }
    }
    else{
        console.warn('It seems that the access token could not be retrieved...');
        DevExpress.ui.notify('It seems that your authenticationd has failed. Please close then open me again', 'error', 1500);
        teamsLoadPanelControl.hide();
        changeButtonStatus(refreshButtonIndex, false);
    }
}

async function loadTeamsFromCacheAsync(){
    console.log('Try loading teams from cache');
    teamsLoadPanelControl.option('message', loadingTeamsMessage);
    teamsLoadPanelControl.show();
    changeButtonStatus(refreshButtonIndex, true);
    changeButtonStatus(shareButtonIndex, true);

    chrome.storage.local.get(
        ['date', 'teams', 'icons'],
        async function(data){
            console.log('Cached data have been loaded');
            if(data && data.date){
                const cacheJsonDateDate = data.date;
                const cacheDate = new Date(cacheJsonDateDate);
                if(cacheHasExpired(cacheDate)){
                    console.log('Cache has expired');
                    loadTeamsFromApiAsync();
                }
                else{
                    //reset teams icons (not saved in cache) and state
                    for (const team of data.teams){
                        team.expanded = false;
                        team.icon = 'icons/teams_32x32.png';
                    }
                    
                    //init the TreeView data store
                    teamsStore = new DevExpress.data.ArrayStore({ key: 'id', data: data.teams });
                    const teamsSource = new DevExpress.data.DataSource({ store: teamsStore, reshapeOnPush: false });
                
                    //load the teams in the TreeView
                    teamsTreeControl.option('dataSource', teamsSource);
                    
                    //icons are not saved in cache...
                    await loadTeamsIconsAsync(data.teams);

                    teamsLoadPanelControl.hide();
                    changeButtonStatus(refreshButtonIndex, false);
                    changeButtonStatus(shareButtonIndex, (teamsTreeControl.selectedItem));
                }
            }
            else{
                console.log('Cache is empty');
                loadTeamsFromApiAsync();
            }
        });
}

async function loadTeamChannelsAsync(team){
    if(requestSettings && team){
        console.log('Loading channels of ' + team.name);

        const channelsRequest = new Request('https://graph.microsoft.com/v1.0/teams/' + team.id + '/channels', requestSettings);
        const channelsResponse = await fetch(channelsRequest, requestSettings)
        .catch(function(error) { console.error(error.message); });

        if(channelsResponse.ok){
            //deserialize returned data
            const channelsData = await channelsResponse.json();
            channelsData.value.sort(compareChannels);
            console.log(channelsData.value.length + ' channels found for ' + team.name);
            
            //prepare data
            var teamChannels = new Array();
            channelsData.value.forEach(channel => {
                teamChannels.push({ id: channel.id, name: channel.displayName, teamId: team.id, teamName: team.name, type: 'channel' });
            });
            //update the team
            teamsStore.push([{ type: 'update', data: { items: teamChannels}, key: team.id }]);
        }
        else{
            console.error('Channels could not be loaded for ' + team.name + '. Status : ' + channelsResponse.status);
            DevExpress.ui.notify('Channels could not be loaded for ' + team.name, 'error', 1500);
        }
    }
}

async function loadTeamsChannelsAsync(teams){
    if(requestSettings){
        teamsLoadPanelControl.option('message', loadingChannelsMessage);
        
        console.log('Loading channels');
        teamsTreeControl.beginUpdate();
        await Promise.all(teams.map(async (team) => {
            await loadTeamChannelsAsync(team);
        }));
        teamsTreeControl.endUpdate();
        console.log('Channels have been loaded');
    }
}

async function loadTeamIconAsync(team){
    if(requestSettings && team){
        console.log('Loading icon of ' + team.name);
        const iconRequest = new Request('https://graph.microsoft.com/v1.0/groups/' + team.id + '/photos/48x48/$value', requestSettings);
        const iconResponse = await fetch(iconRequest, requestSettings)
        .catch(function(error) { console.error(error.message); });

        if(iconResponse.ok){
            //deserialize returned data
            const iconBlob = await iconResponse.blob();
            const iconUrl = URL.createObjectURL(iconBlob);
            //update the team
            teamsStore.push([{ type: 'update', data: { icon: iconUrl}, key: team.id }]);
        }
        else{
            console.log('Icon could not be loaded for ' + team.name + '. Status' + iconResponse.status);
        }
    }
}

async function loadTeamsIconsAsync(teams){
    if(requestSettings){
        console.log('Loading icon of each team');
        await Promise.all(teams.map(async (team) => {
            await loadTeamIconAsync(team);
          }));
          console.log('Icons have been loaded');
    }
}


async function shareAsync(){
    if(!shareButtonEnabled){
        return;
    }
    if(requestSettings){
        if(teamsTreeControl.selectedItem){
            console.log('Sharing page with ' + teamsTreeControl.selectedItem.name + '(' + teamsTreeControl.selectedItem.teamName + ')');
            teamsLoadPanelControl.option('message', 'Sharing...');
            changeButtonStatus(shareButtonIndex, true);
            changeButtonStatus(refreshButtonIndex, true);
            teamsLoadPanelControl.show();

            //create message content
            var message = createMessage();
            var json = JSON.stringify(message);
            //send message to Teams
            var messageRequestSettings = { method: 'POST', headers: requestSettings.headers, body: json };
            const messageRequest = new Request('https://graph.microsoft.com/beta/teams/' + teamsTreeControl.selectedItem.teamId + '/channels/' + teamsTreeControl.selectedItem.id + '/messages', messageRequestSettings);
            const messageResponse = await fetch(messageRequest, messageRequestSettings)
            .catch(function(error) { console.error('ERROR: ' + error.message); });

            if(messageResponse.ok){
                const responseData = await messageResponse.json();
                console.log('The message has been sent to Teams (message id : ' + responseData.id + ')');
                teamsLoadPanelControl.hide();
                DevExpress.ui.notify('The page has been shared !', 'success', 1000);
                close();
            }
            else{
                DevExpress.ui.notify('Teams has rejected the request', 'error', 1500);
                console.error('Teams has rejected the request. Status ' + messageResponse.status);
                changeButtonStatus(shareButtonIndex, false);
                changeButtonStatus(refreshButtonIndex, false);
                teamsLoadPanelControl.hide();
            }
        }
        else{
            DevExpress.ui.notify('Please select a channel', 'warning', 1500);
        }
    }
}

//#endregion

//#region People Picker

async function getEmployeePhotoAsync(id){
    if(requestSettings && id){
        console.log('Get photo of employee ' + id); 
        const photoRequest = new Request('https://graph.microsoft.com/beta/users/' + id + '/photos/48x48/$value', requestSettings);
        const photoResponse = await fetch(photoRequest, requestSettings)
        .catch(function(error) { console.error(error.message); });

        if(photoResponse.ok){
            console.log('Photo has been found for ' + id);
            //deserialize the returned data
            const photoBlob = await photoResponse.blob();
            return URL.createObjectURL(photoBlob);
        }
        else if(photoResponse.status == 404){
            console.log('Photo not found for ' + id); 
            return 'icons/Unknown_64x64.png';
        }
        else{
            console.warn('The photo could not be retrieved for ' + id + ' .Status : ' + photoResponse.status);
            return 'icons/Unknown_64x64.png';
        }
    }
}

function searchEmployees(){
    if (searchEmployeesTimeout) {
        window.clearTimeout(searchEmployeesTimeout);
    }
    searchEmployeesTimeout = window.setTimeout(searchEmployeesAsync, 1000);
}

async function searchEmployeesAsync() {
    tagsBoxControl.close();
    tagsBoxControl.option('dataSource', new Array());

    var searchFilter = tagsBoxControl.option('text');
    if (searchFilter.length >= 3) {
        tagsBoxControl.option('readOnly', true);
        tagsLoadIndicatorControl.option('visible', true);

        console.log('Search people with filter "' + searchFilter + '" (startsWith, attributes : mail, displayName, givenName, surname)');
        //construct request
        const query = "?$filter=startswith(mail, '" + searchFilter + "')"
                    + " or startswith(displayName, '" + searchFilter + "')"
                    + " or startswith(givenName, '" + searchFilter + "')"
                    + " or startswith(surname, '" + searchFilter + "')";
        const requestUrl = 'https://graph.microsoft.com/v1.0/users/' + query + '&$select=id,mail,displayName,jobTitle&$top=5';

        //send request to Graph Api
        const peopleRequest = new Request(requestUrl, requestSettings);
        const peopleResponse = await fetch(peopleRequest, requestSettings)
        .catch(function(error) { console.error(error.message); });

        if(peopleResponse.ok){
            //deserialize returned data
            const peopleData = await peopleResponse.json();
            if(peopleData.value){
                console.log(peopleData.value.length + ' people found');
                var employees = new Array();
                await Promise.all(peopleData.value.map(async (employee) => {
                    var photo = await getEmployeePhotoAsync(employee.id);
                    employees.push({
                        id: employee.id, 
                        mail: employee.mail, 
                        displayName: employee.displayName, 
                        jobTitle: employee.jobTitle, 
                        photo: photo
                    });
                  }));
                
                //set then show list for user
                tagsBoxControl.option('dataSource', employees);
                console.log('People picker updated');
                tagsBoxControl.open();
            }
            else{
                console.log('No people found');
            }
        }
        else{
            console.error('People could not be found. Status : ' + peopleResponse.status);
        }

        tagsBoxControl.option('readOnly', false);
        tagsLoadIndicatorControl.option('visible', false);
    }
}

//#endregion


function init(){
    chrome.tabs.query({
        active: true,
        currentWindow: true
    },
    function (tabs){
        currentTab = tabs[0];
        var title = document.getElementById('title');
        if(title){
            title.innerHTML = 'Share <i>' + truncateText(currentTab.title) + '</i>';
        }
        console.log('Share - Title : ' + currentTab.title);
        console.log('Share - Url ' + currentTab.url);

        const url = new URL(currentTab.url);
        chrome.tabs.executeScript(currentTab.id, { code: tabMetaScript }, function(result) {
            if(result){
                const { title, description, image } = result[0];
                metaTitle = title;
                metaDescription = description;
                console.log('Card - Title : ' + title);
                console.log('Card - Description : ' + description);
                
                if(image) {
                    metaImage = image;
                    if(Object.prototype.toString.call(metaImage) == '[object String]'){
                        if(metaImage.startsWith("/")){
                            var currentOrigin = (new URL(currentTab.url)).origin;
                            if(currentOrigin.endsWith('/')){
                                metaImage = currentOrigin + metaImage.substring(1);
                            }
                            else{
                                metaImage = currentOrigin + metaImage;
                            }
                        }
                    }
                    console.log('Card - Image : ' + metaImage);
                }
                else{
                    console.log('Card - No image found');
                }
            }

            getTeams(false);
          });
    });
}


//Initialize controls and process
$(function(){
    teamsTreeControl = $('#teamsContainer').dxTreeView({
        dataStructure: 'tree',
        displayExpr: 'name',
        expandEvent: 'click',
        height: 350,
        noDataText: '',
        onItemClick: async function(e) {
            var item = e.itemData;
            if(item.type == 'channel'){
                selectedChannel = e.itemData;
                e.component.selectedItem = e.itemData;
                changeButtonStatus(shareButtonIndex, false);
            }
            else if(item.type == 'team'){
                selectedChannel = null;
                e.component.selectedItem = null;
                teamElementToScroll = e.itemElement;
                changeButtonStatus(shareButtonIndex, true);
            }
        },
        rootValue: '',
        searchEnabled: true,
        searchExpr: ['name', 'teamName'],
        searchMode: 'contains',
        searchEditorOptions: {
            placeholder: 'Search a team...'
        },
        selectByClick: true,
        selectionMode: 'single',
        selectNodesRecursive: false
    }).dxTreeView("instance");
    
    teamsLoadPanelControl = $('#loadPanelContainer').dxLoadPanel({
        closeOnOutsideClick: false,
        message: loadingTeamsMessage,
        position: { my: 'center', at: 'center', of: '#teamsContainer' }
    }).dxLoadPanel('instance');

    toolbarControl = $('#toolbarContainer').dxToolbar({
        items:[
        {
            location: 'after',
            locateInMenu: 'always',
            widget: 'dxButton',
            options: {
                disabled: true,
                icon: 'refresh',
                id: 'btnRefresh',
                onClick: () => { getTeams(true) },
                text: 'Refresh my teams',
                type: 'normal'
            },
        }]
    }).dxToolbar('instance');

    txtSubjectControl = $('#txtSubject').dxTextBox({
        placeholder: 'Add a subject ?',
        showClearButton: true,
        width: 360
    }).dxTextBox('instance');

    txtMessageControl = $('#txtMessage').dxTextArea({
        placeholder: 'Add a message ?',
        showClearButton: true,
        width: 360
    }).dxTextArea('instance');

    isImportantControl = $('#chkImportant').dxSwitch({
        switchedOffText: 'No',
        switchedOnText: 'Yes'
    }).dxSwitch('instance');
    $("#chkImportant").find("div.dx-switch-handle").addClass("chkImportant").removeClass(".dx-switch-on-value");

    mentionTeamControl = $('#chkMentionTeam').dxSwitch({
        switchedOffText: 'No',
        switchedOnText: 'Yes'
    }).dxSwitch('instance');

    mentionChannelControl = $('#chkMentionChannel').dxSwitch({
        switchedOffText: 'No',
        switchedOnText: 'Yes'
    }).dxSwitch('instance');

    tagsBoxControl = $('#tagsMentionPeople').dxTagBox({
        acceptCustomValue: true,
        displayExpr: 'displayName',
        itemTemplate: function (itemData, itemIndex, itemElement) {
            return $('<div />').append(
                $('<div />').addClass('employeeContainer').append(
                    $('<div />').addClass('employeePhotoContainer').append($('<img />').attr('src', itemData.photo).addClass('employeePhoto')),
                    $('<div />').addClass('employeeDataContainer').append(
                        $('<div />').addClass('employeeDisplayName').text(itemData.displayName),
                        $('<div />').addClass('employeeJobTitle').text(itemData.jobTitle),
                        $('<div />').addClass('employeeMail').text(itemData.mail?.toLowerCase()))
                )
            );
        },
        noDataText: 'We did not find anyone :(',
        onKeyUp: searchEmployees,
        openOnFieldClick: false,
        placeholder: 'Mention your colleagues...',
        searchEnabled: false,
        showClearButton: true,
        showDataBeforeSearch: false,
        tagTemplate: function(tagData) {
            return $("<div />")
                .addClass("dx-tag-content")
                .append(
                    $('<div />').addClass('employeeSmallPhotoContainer').append($('<img />').attr('src', tagData.photo).addClass('employeeSmallPhoto')),
                    $('<div />').css('float', 'left').append($("<span />").text(tagData.displayName)),
                    $("<div />").addClass("dx-tag-remove-button")
                );
        }
    }).dxTagBox('instance');

    tagsLoadIndicatorControl = $('#tagsLoadIndicator').dxLoadIndicator({
        height: 36,
        visible: false,
        width: 36
    }).dxLoadIndicator('instance');

    $('#dialShare').dxSpeedDialAction({
        hint: 'Share',
        icon: 'share',
        onClick: shareAsync
    });

    
    teamsLoadPanelControl.show();
    init();
});