// Test function to authorize scopes
function authorizeScopes() {
  try {
    var test1 = Session.getActiveUser().getEmail();
    Logger.log('User email: ' + test1);
    var test2 = DriveApp.getRootFolder().getName();
    Logger.log('Drive root: ' + test2);
    var optionalArgs = {
      'q': "mimeType='application/vnd.google-apps.folder' and trashed=false",
      'maxResults': 1
    };
    var test3 = Drive.Files.list(optionalArgs);
    Logger.log('Drive API works: ' + test3.items.length);
    return 'All scopes authorized successfully!';
  } catch (error) {
    Logger.log('Authorization error: ' + error);
    return 'Error: ' + error;
  }
}

function getFavoriteFolders() {
  try {
    var userProperties = PropertiesService.getUserProperties();
    var favoritesJson = userProperties.getProperty('favoriteFolders');
    return favoritesJson ? JSON.parse(favoritesJson) : [];
  } catch (error) {
    console.error('Error getting favorite folders:', error);
    return [];
  }
}

function saveFavoriteFolders(favorites) {
  try {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('favoriteFolders', JSON.stringify(favorites));
    return true;
  } catch (error) {
    console.error('Error saving favorite folders:', error);
    return false;
  }
}

function addToFavorites(e) {
  try {
    var folderId = e.parameters.folderId;
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    var folderName = 'Unknown Folder';
    try {
      folderName = DriveApp.getFolderById(folderId).getName();
    } catch (err) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Invalid folder ID or no access'))
        .build();
    }
    
    var favorites = getFavoriteFolders();
    
    for (var i = 0; i < favorites.length; i++) {
      if (favorites[i].id === folderId) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification().setText('Already in quick access'))
          .build();
      }
    }
    
    if (favorites.length >= 10) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Maximum 10 folders. Remove one first.'))
        .build();
    }
    
    favorites.push({ id: folderId, name: folderName });
    saveFavoriteFolders(favorites);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Added to Quick Access: ' + folderName))
      .build();
  } catch (error) {
    console.error('Error in addToFavorites:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Unable to add folder'))
      .build();
  }
}

function showFavoritesManager(e) {
  try {
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    var favorites = getFavoriteFolders();
    
    var header = CardService.newCardHeader()
      .setTitle('Quick Access Folders')
      .setSubtitle(favorites.length + ' of 10 folders saved')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/star_black_48dp.png');
    
    var section = CardService.newCardSection();
    
    if (favorites.length > 0) {
      for (var i = 0; i < favorites.length; i++) {
        var fav = favorites[i];
        var removeAction = CardService.newAction()
          .setFunctionName('removeFavorite')
          .setParameters({
            'folderId': fav.id,
            'messageId': messageId,
            'numAtts': numAtts
          });
        
        var favWidget = CardService.newDecoratedText()
          .setText('📁 ' + fav.name)
          .setButton(CardService.newTextButton()
            .setText('✕ Remove')
            .setOnClickAction(removeAction));
        section.addWidget(favWidget);
      }
      
      section.addWidget(CardService.newTextParagraph()
        .setText('<font color="#666666">' + favorites.length + ' of 10 folders</font>'));
    } else {
      section.addWidget(CardService.newTextParagraph()
        .setText('No quick access folders yet. When browsing, click "⭐ Quick Access" on any folder.'));
    }
    
    section.addWidget(CardService.newDivider());
    
    var folderIdInput = CardService.newTextInput()
      .setFieldName('newFavFolderId')
      .setTitle('Add by Folder ID')
      .setHint('Paste folder ID from Drive URL');
    section.addWidget(folderIdInput);
    
    var addAction = CardService.newAction()
      .setFunctionName('addFavoriteFromManager')
      .setParameters({ 'messageId': messageId, 'numAtts': numAtts });
    var addButton = CardService.newTextButton()
      .setText('➕ Add Folder')
      .setOnClickAction(addAction)
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED);
    section.addWidget(addButton);
    
    var card = CardService.newCardBuilder()
      .setHeader(header)
      .addSection(section)
      .setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
          .setText('🏠 Done')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('returnToHome')
            .setParameters({ 'messageId': messageId, 'numAtts': numAtts }))));
    
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(card.build()))
      .build();
  } catch (error) {
    console.error('Error in showFavoritesManager:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error loading manager'))
      .build();
  }
}

function addFavoriteFromManager(e) {
  try {
    var folderId = e.formInputs.newFavFolderId ? e.formInputs.newFavFolderId[0].trim() : '';
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    if (!folderId) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Enter a folder ID'))
        .build();
    }
    
    if (folderId.length < 20 || folderId.includes(' ')) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Invalid folder ID format'))
        .build();
    }
    
    var favorites = getFavoriteFolders();
    
    if (favorites.length >= 10) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Maximum 10 folders. Remove one first.'))
        .build();
    }
    
    var folderName = 'Unknown Folder';
    try {
      folderName = DriveApp.getFolderById(folderId).getName();
    } catch (err) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Invalid folder ID or no access'))
        .build();
    }
    
    for (var i = 0; i < favorites.length; i++) {
      if (favorites[i].id === folderId) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification().setText('Already in quick access'))
          .build();
      }
    }
    
    favorites.push({ id: folderId, name: folderName });
    saveFavoriteFolders(favorites);
    
    return showFavoritesManager({
      parameters: { messageId: messageId, numAtts: numAtts }
    });
  } catch (error) {
    console.error('Error in addFavoriteFromManager:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Unable to add folder'))
      .build();
  }
}

function removeFavorite(e) {
  try {
    var folderId = e.parameters.folderId;
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    var favorites = getFavoriteFolders();
    var newFavorites = [];
    
    for (var i = 0; i < favorites.length; i++) {
      if (favorites[i].id !== folderId) {
        newFavorites.push(favorites[i]);
      }
    }
    
    saveFavoriteFolders(newFavorites);
    
    return showFavoritesManager({
      parameters: { messageId: messageId, numAtts: numAtts }
    });
  } catch (error) {
    console.error('Error in removeFavorite:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error removing folder'))
      .build();
  }
}

function onGmailMessage(e) {
  var messageId = e.gmail.messageId;
  var accessToken = e.gmail.accessToken;
  
  GmailApp.setCurrentMessageAccessToken(accessToken);
  
  var message = GmailApp.getMessageById(messageId);
  var attachments = message.getAttachments();
  
  if (attachments.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('No attachments in this email'))
      .build();
  }
  
  var cardHeader = CardService.newCardHeader()
    .setTitle('Save Attachments')
    .setSubtitle(attachments.length + ' file' + (attachments.length > 1 ? 's' : '') + ' • Choose destination')
    .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/cloud_upload_black_48dp.png');
  
  var quickSection = CardService.newCardSection();
  
  var favorites = getFavoriteFolders();
  if (favorites && favorites.length > 0) {
    quickSection.setHeader('<b>Quick Access</b>');
    
    for (var f = 0; f < Math.min(favorites.length, 5); f++) {
      var fav = favorites[f];
      var favAction = CardService.newAction()
        .setFunctionName('selectFolder')
        .setParameters({
          'selectedFolderId': fav.id,
          'messageId': messageId,
          'numAtts': attachments.length.toString()
        });
      var favButton = CardService.newTextButton()
        .setText(fav.name)
        .setOnClickAction(favAction);
      quickSection.addWidget(favButton);
    }
    
    if (favorites.length > 5) {
      quickSection.addWidget(CardService.newTextParagraph()
        .setText('<font color="#666666">+' + (favorites.length - 5) + ' more in manager</font>'));
    }
  }
  
  var manageFavAction = CardService.newAction()
    .setFunctionName('showFavoritesManager')
    .setParameters({ 'messageId': messageId, 'numAtts': attachments.length.toString() });
  var manageFavButton = CardService.newTextButton()
    .setText(favorites.length > 0 ? 'Manage Quick Access' : 'Set Up Quick Access')
    .setOnClickAction(manageFavAction);
  quickSection.addWidget(manageFavButton);
  
  var browseSection = CardService.newCardSection()
    .setHeader('<b>Browse Folders</b>');
  
  var searchAction = CardService.newAction()
    .setFunctionName('buildSearchResults')
    .setParameters({ 'messageId': messageId, 'numAtts': attachments.length.toString() });
  
  var searchDecoratedText = CardService.newDecoratedText()
    .setTopLabel('Search for folder')
    .setText('')
    .setButton(CardService.newTextButton()
      .setText('Search')
      .setOnClickAction(searchAction));
  
  var searchInput = CardService.newTextInput()
    .setFieldName('searchQuery')
    .setHint('Type folder name...');
  
  browseSection.addWidget(searchInput);
  browseSection.addWidget(searchDecoratedText);
  browseSection.addWidget(CardService.newDivider());
  
  var browseMyAction = CardService.newAction()
    .setFunctionName('buildFolderBrowser')
    .setParameters({ 
      'currentFolderId': DriveApp.getRootFolder().getId(), 
      'parentFolderId': '', 
      'grandParentId': '',
      'currentPath': 'My Drive',
      'messageId': messageId, 
      'numAtts': attachments.length.toString() 
    });
  var browseMyButton = CardService.newTextButton()
    .setText('My Drive')
    .setOnClickAction(browseMyAction);
  browseSection.addWidget(browseMyButton);
  
  var browseStarredAction = CardService.newAction()
    .setFunctionName('buildStarredBrowser')
    .setParameters({ 
      'parentFolderId': '', 
      'grandParentId': '',
      'currentPath': 'Starred',
      'messageId': messageId, 
      'numAtts': attachments.length.toString() 
    });
  var browseStarredButton = CardService.newTextButton()
    .setText('Starred')
    .setOnClickAction(browseStarredAction);
  browseSection.addWidget(browseStarredButton);
  
  var advancedSection = CardService.newCardSection()
    .setHeader('<b>More Options</b>')
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0);
  
  var recentFolders = getRecentFolders(5);
  if (recentFolders && recentFolders.length > 0) {
    var folderSelection = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle('Recent folders')
      .setFieldName('folderId');
    
    for (var k = 0; k < recentFolders.length; k++) {
      var folder = recentFolders[k];
      folderSelection.addItem(folder.title, folder.id, false);
    }
    advancedSection.addWidget(folderSelection);
  }
  
  var folderInput = CardService.newTextInput()
    .setFieldName('customFolderId')
    .setTitle('Folder ID')
    .setHint('Paste folder ID...');
  advancedSection.addWidget(folderInput);
  
  var createFolderAction = CardService.newAction()
    .setFunctionName('showCreateFolderDialog')
    .setParameters({ 'messageId': messageId, 'numAtts': attachments.length.toString() });
  var createFolderButton = CardService.newTextButton()
    .setText('➕ Create New Folder')
    .setOnClickAction(createFolderAction);
  advancedSection.addWidget(createFolderButton);
  
  var attachmentSection = CardService.newCardSection()
    .setHeader('<b>Filename Options</b>');
  
  var useOriginalSwitch = CardService.newSwitch()
    .setFieldName('addTimestamp')
    .setValue('on')
    .setSelected(true);
  var switchDecoratedText = CardService.newDecoratedText()
    .setTopLabel('Rename files')
    .setText('Add timestamp: invoice.pdf → invoice_2025-10-04T14-30.pdf')
    .setSwitchControl(useOriginalSwitch);
  attachmentSection.addWidget(switchDecoratedText);
  
  attachmentSection.addWidget(CardService.newDivider());
  
  var totalSize = 0;
  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    var sizeInMB = (att.getSize() / (1024 * 1024)).toFixed(2);
    totalSize += att.getSize();
    var originalName = att.getName();
    var input = CardService.newTextInput()
      .setFieldName('newName' + i)
      .setTitle(originalName + ' (' + sizeInMB + ' MB)')
      .setValue(originalName)
      .setHint('Leave blank to skip');
    attachmentSection.addWidget(input);
  }
  
  if (totalSize > 10 * 1024 * 1024) {
    var warningText = CardService.newTextParagraph()
      .setText('<font color="#d93025">⚠ Large files detected. Saving may take longer.</font>');
    attachmentSection.addWidget(warningText);
  }
  
  var card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .addSection(quickSection)
    .addSection(browseSection)
    .addSection(advancedSection)
    .addSection(attachmentSection)
    .setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('Save to Drive')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('saveAttachments')
          .setParameters({ 'messageId': messageId, 'numAtts': attachments.length.toString() }))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)));
  
  return card.build();
}

function getRecentFolders(limit) {
  try {
    var cacheKey = 'recentFolders_' + limit;
    var cache = CacheService.getUserCache();
    var cached = cache.get(cacheKey);
    
    if (cached) {
      return JSON.parse(cached);
    }
    
    var optionalArgs = {
      'q': "mimeType='application/vnd.google-apps.folder' and trashed=false",
      'orderBy': 'modifiedDate desc',
      'fields': 'items(id,title,modifiedDate)',
      'maxResults': limit || 10
    };
    var response = Drive.Files.list(optionalArgs);
    var items = response.items || [];
    
    cache.put(cacheKey, JSON.stringify(items), 300);
    
    return items;
  } catch (error) {
    console.error('Error fetching folders:', error);
    return [];
  }
}

function buildSearchResults(e) {
  try {
    var searchQuery = e.formInputs.searchQuery ? e.formInputs.searchQuery[0].trim() : '';
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    if (!searchQuery) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Enter a folder name'))
        .build();
    }
    
    var formattedQuery = "title contains '" + searchQuery.replace(/'/g, "\\'") + "'";
    var searchedFolders = DriveApp.searchFolders(formattedQuery);
    var searchList = [];
    
    var count = 0;
    while (searchedFolders.hasNext() && count < 15) {
      var folder = searchedFolders.next();
      searchList.push({
        id: folder.getId(),
        name: folder.getName()
      });
      count++;
    }
    
    var searchHeader = CardService.newCardHeader()
      .setTitle('Search Results')
      .setSubtitle('"' + searchQuery + '" • ' + count + ' folder' + (count !== 1 ? 's' : '') + ' found')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/search_black_48dp.png');
    
    var searchSection = CardService.newCardSection();
    
    if (count === 0) {
      searchSection.addWidget(CardService.newTextParagraph()
        .setText('No folders found. Try a different search term.'));
    } else {
      for (var n = 0; n < searchList.length; n++) {
        var result = searchList[n];
        
        var selectAction = CardService.newAction()
          .setFunctionName('selectFolder')
          .setParameters({ 
            'selectedFolderId': result.id, 
            'messageId': messageId, 
            'numAtts': numAtts 
          });
        
        var browseAction = CardService.newAction()
          .setFunctionName('buildFolderBrowser')
          .setParameters({ 
            'currentFolderId': result.id, 
            'parentFolderId': '', 
            'grandParentId': '',
            'currentPath': result.name,
            'messageId': messageId, 
            'numAtts': numAtts 
          });
        
        var resultWidget = CardService.newDecoratedText()
          .setText(result.name)
          .setButton(CardService.newTextButton()
            .setText('✓ Select')
            .setOnClickAction(selectAction))
          .setButton(CardService.newTextButton()
            .setText('→ Browse')
            .setOnClickAction(browseAction));
        searchSection.addWidget(resultWidget);
      }
    }
    
    var searchCard = CardService.newCardBuilder()
      .setHeader(searchHeader)
      .addSection(searchSection)
      .setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
          .setText('Home')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('returnToHome')
            .setParameters({ 'messageId': messageId, 'numAtts': numAtts }))));
    
    var navigation = CardService.newNavigation().pushCard(searchCard.build());
    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .build();
  } catch (error) {
    console.error('Error in buildSearchResults:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Search error'))
      .build();
  }
}

function buildFolderBrowser(e) {
  try {
    var currentFolderId = e.parameters.currentFolderId;
    var parentFolderId = e.parameters.parentFolderId || '';
    var grandParentId = e.parameters.grandParentId || '';
    var currentPath = e.parameters.currentPath;
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    var optionalArgs = {
      'q': "'" + currentFolderId + "' in parents and (mimeType='application/vnd.google-apps.folder' or mimeType='application/vnd.google-apps.shortcut') and trashed=false",
      'fields': 'items(id,title,mimeType,shortcutDetails(targetId,targetMimeType))',
      'pageSize': 20
    };
    var response = Drive.Files.list(optionalArgs);
    var subfolderList = [];
    
    for (var m = 0; m < response.items.length; m++) {
      var item = response.items[m];
      var displayName = item.title;
      var targetId = item.id;
      if (item.mimeType === 'application/vnd.google-apps.shortcut') {
        if (item.shortcutDetails && item.shortcutDetails.targetMimeType === 'application/vnd.google-apps.folder') {
          displayName = item.title + ' (shared)';
          targetId = item.shortcutDetails.targetId;
        } else {
          continue;
        }
      }
      subfolderList.push({ id: targetId, name: displayName });
    }
    
    var browserHeader = CardService.newCardHeader()
      .setTitle(DriveApp.getFolderById(currentFolderId).getName())
      .setSubtitle(currentPath || 'My Drive')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/folder_open_black_48dp.png');
    
    var browserSection = CardService.newCardSection();
    
    if (parentFolderId) {
      var backAction = CardService.newAction()
        .setFunctionName('buildFolderBrowser')
        .setParameters({ 
          'currentFolderId': parentFolderId, 
          'parentFolderId': grandParentId, 
          'grandParentId': '',
          'currentPath': currentPath ? currentPath.substring(0, currentPath.lastIndexOf(' / ')) : 'My Drive',
          'messageId': messageId, 
          'numAtts': numAtts 
        });
      var backButton = CardService.newTextButton()
        .setText('⬆ Up')
        .setOnClickAction(backAction);
      browserSection.addWidget(backButton);
    }
    
    if (subfolderList.length === 0) {
      browserSection.addWidget(CardService.newTextParagraph()
        .setText('<font color="#666666">No subfolders</font>'));
    }
    
    for (var m = 0; m < subfolderList.length; m++) {
      var sub = subfolderList[m];
      var newPath = currentPath ? currentPath + ' / ' + sub.name : 'My Drive / ' + sub.name;
      var subAction = CardService.newAction()
        .setFunctionName('navigateToSubfolder')
        .setParameters({ 
          'subFolderId': sub.id, 
          'parentFolderId': currentFolderId, 
          'grandParentId': parentFolderId,
          'currentPath': newPath,
          'messageId': messageId, 
          'numAtts': numAtts 
        });
      var subButton = CardService.newTextButton()
        .setText(sub.name)
        .setOnClickAction(subAction);
      browserSection.addWidget(subButton);
    }
    
    var addFavAction = CardService.newAction()
      .setFunctionName('addToFavorites')
      .setParameters({
        'folderId': currentFolderId,
        'messageId': messageId,
        'numAtts': numAtts
      });
    
    var browserCard = CardService.newCardBuilder()
      .setHeader(browserHeader)
      .addSection(browserSection)
      .setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
          .setText('✓ Select This Folder')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('selectFolder')
            .setParameters({ 
              'selectedFolderId': currentFolderId, 
              'messageId': messageId, 
              'numAtts': numAtts 
            }))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .setSecondaryButton(CardService.newTextButton()
          .setText('⭐ Quick Access')
          .setOnClickAction(addFavAction)));
    
    var navigation = CardService.newNavigation().pushCard(browserCard.build());
    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .build();
  } catch (error) {
    console.error('Error in buildFolderBrowser:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error browsing'))
      .build();
  }
}

function navigateToSubfolder(e) {
  try {
    return buildFolderBrowser({ 
      parameters: { 
        currentFolderId: e.parameters.subFolderId, 
        parentFolderId: e.parameters.parentFolderId,
        grandParentId: e.parameters.grandParentId,
        currentPath: e.parameters.currentPath,
        messageId: e.parameters.messageId, 
        numAtts: e.parameters.numAtts 
      }
    });
  } catch (error) {
    console.error('Error in navigateToSubfolder:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Navigation error'))
      .build();
  }
}

function buildStarredBrowser(e) {
  try {
    var parentFolderId = e.parameters.parentFolderId || '';
    var grandParentId = e.parameters.grandParentId || '';
    var currentPath = e.parameters.currentPath;
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    var starredFolders;
    if (!parentFolderId) {
      starredFolders = DriveApp.searchFolders('starred = true');
    } else {
      starredFolders = DriveApp.getFolderById(parentFolderId).getFolders();
    }
    
    var subfolderList = [];
    var count = 0;
    while (starredFolders.hasNext() && count < 20) {
      var folder = starredFolders.next();
      subfolderList.push({
        id: folder.getId(),
        name: folder.getName()
      });
      count++;
    }
    
    var browserHeader = CardService.newCardHeader()
      .setTitle(parentFolderId ? DriveApp.getFolderById(parentFolderId).getName() : 'Starred Folders')
      .setSubtitle(currentPath || 'Starred folders')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/star_black_48dp.png');
    
    var browserSection = CardService.newCardSection();
    
    if (parentFolderId) {
      var backAction = CardService.newAction()
        .setFunctionName('buildStarredBrowser')
        .setParameters({ 
          'parentFolderId': grandParentId, 
          'grandParentId': '',
          'currentPath': currentPath ? currentPath.substring(0, currentPath.lastIndexOf(' / ')) : 'Starred',
          'messageId': messageId, 
          'numAtts': numAtts 
        });
      var backButton = CardService.newTextButton()
        .setText('Up')
        .setOnClickAction(backAction);
      browserSection.addWidget(backButton);
    }
    
    if (count === 0 && !parentFolderId) {
      browserSection.addWidget(CardService.newTextParagraph()
        .setText('<font color="#666666">No starred folders. Star a folder in Drive to see it here.</font>'));
    } else if (count === 0) {
      browserSection.addWidget(CardService.newTextParagraph()
        .setText('<font color="#666666">No subfolders</font>'));
    }
    
    for (var m = 0; m < subfolderList.length; m++) {
      var sub = subfolderList[m];
      var newPath = currentPath ? currentPath + ' / ' + sub.name : 'Starred / ' + sub.name;
      var subAction = CardService.newAction()
        .setFunctionName('navigateToStarredSubfolder')
        .setParameters({ 
          'subFolderId': sub.id, 
          'parentFolderId': parentFolderId || sub.id, 
          'grandParentId': parentFolderId,
          'currentPath': newPath,
          'messageId': messageId, 
          'numAtts': numAtts 
        });
      var subButton = CardService.newTextButton()
        .setText(sub.name)
        .setOnClickAction(subAction);
      browserSection.addWidget(subButton);
    }
    
    var currentStarredId = parentFolderId || (subfolderList.length > 0 ? subfolderList[0].id : null);
    
    var browserCard = CardService.newCardBuilder()
      .setHeader(browserHeader)
      .addSection(browserSection);
    
    if (currentStarredId) {
      var addFavAction = CardService.newAction()
        .setFunctionName('addToFavorites')
        .setParameters({
          'folderId': currentStarredId,
          'messageId': messageId,
          'numAtts': numAtts
        });
      
      browserCard.setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
          .setText('Select This Folder')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('selectFolder')
            .setParameters({ 
              'selectedFolderId': currentStarredId, 
              'messageId': messageId, 
              'numAtts': numAtts 
            }))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .setSecondaryButton(CardService.newTextButton()
          .setText('Add to Quick Access')
          .setOnClickAction(addFavAction)));
    }
    
    var navigation = CardService.newNavigation().pushCard(browserCard.build());
    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .build();
  } catch (error) {
    console.error('Error in buildStarredBrowser:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error browsing starred'))
      .build();
  }
}

function navigateToStarredSubfolder(e) {
  try {
    return buildStarredBrowser({ 
      parameters: { 
        parentFolderId: e.parameters.subFolderId, 
        grandParentId: e.parameters.parentFolderId,
        currentPath: e.parameters.currentPath,
        messageId: e.parameters.messageId, 
        numAtts: e.parameters.numAtts 
      }
    });
  } catch (error) {
    console.error('Error in navigateToStarredSubfolder:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Navigation error'))
      .build();
  }
}

function showCreateFolderDialog(e) {
  try {
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    var dialogHeader = CardService.newCardHeader()
      .setTitle('Create New Folder')
      .setSubtitle('Choose location and name')
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/create_new_folder_black_48dp.png');
    
    var dialogSection = CardService.newCardSection();
    
    var folderNameInput = CardService.newTextInput()
      .setFieldName('newFolderName')
      .setTitle('Folder name')
      .setHint('Enter folder name');
    dialogSection.addWidget(folderNameInput);
    
    var parentFolderInput = CardService.newTextInput()
      .setFieldName('parentFolderId')
      .setTitle('Parent folder ID (optional)')
      .setHint('Leave blank for My Drive root')
      .setValue(DriveApp.getRootFolder().getId());
    dialogSection.addWidget(parentFolderInput);
    
    var dialogCard = CardService.newCardBuilder()
      .setHeader(dialogHeader)
      .addSection(dialogSection)
      .setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
          .setText('Create')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('createNewFolder')
            .setParameters({ 'messageId': messageId, 'numAtts': numAtts }))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .setSecondaryButton(CardService.newTextButton()
          .setText('Cancel')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('returnToHome')
            .setParameters({ 'messageId': messageId, 'numAtts': numAtts }))));
    
    var navigation = CardService.newNavigation().pushCard(dialogCard.build());
    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .build();
  } catch (error) {
    console.error('Error in showCreateFolderDialog:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error opening dialog'))
      .build();
  }
}

function createNewFolder(e) {
  try {
    var folderName = e.formInputs.newFolderName ? e.formInputs.newFolderName[0].trim() : '';
    var parentFolderId = e.formInputs.parentFolderId ? e.formInputs.parentFolderId[0] : DriveApp.getRootFolder().getId();
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    if (!folderName) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Enter a folder name'))
        .build();
    }
    
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var newFolder = parentFolder.createFolder(folderName);
    
    return selectFolder({
      parameters: {
        selectedFolderId: newFolder.getId(),
        messageId: messageId,
        numAtts: numAtts
      },
      gmail: { accessToken: e.gmail.accessToken }
    });
  } catch (error) {
    console.error('Error in createNewFolder:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Error creating folder'))
      .build();
  }
}

function selectFolder(e) {
  try {
    var selectedFolderId = e.parameters.selectedFolderId;
    var messageId = e.parameters.messageId;
    var numAtts = parseInt(e.parameters.numAtts);
    var accessToken = e.gmail ? e.gmail.accessToken : null;
    
    if (accessToken) {
      GmailApp.setCurrentMessageAccessToken(accessToken);
    }
    
    var message = GmailApp.getMessageById(messageId);
    var attachments = message.getAttachments();
    
    if (attachments.length === 0) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('No attachments found'))
        .build();
    }
    
    var folderName = 'Selected Folder';
    try {
      folderName = DriveApp.getFolderById(selectedFolderId).getName();
    } catch (err) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Cannot access folder'))
        .build();
    }
    
    var cardHeader = CardService.newCardHeader()
      .setTitle(folderName)
      .setSubtitle('Ready to save ' + attachments.length + ' file' + (attachments.length > 1 ? 's' : ''))
      .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/folder_black_48dp.png');
    
    var folderSection = CardService.newCardSection();
    
    var folderInfo = CardService.newDecoratedText()
      .setTopLabel('Save to')
      .setText(folderName)
      .setButton(CardService.newTextButton()
        .setText('Change')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('returnToHome')
          .setParameters({ 'messageId': messageId, 'numAtts': numAtts.toString() })));
    folderSection.addWidget(folderInfo);
    
    var folderInput = CardService.newTextInput()
      .setFieldName('customFolderId')
      .setValue(selectedFolderId)
      .setHint('Folder ID');
    folderSection.addWidget(folderInput);
    
    var attachmentSection = CardService.newCardSection()
      .setHeader('<b>Files</b>');
    
    var useOriginalSwitch = CardService.newSwitch()
      .setFieldName('addTimestamp')
      .setValue('on')
      .setSelected(true);
    var switchText = CardService.newDecoratedText()
      .setTopLabel('Rename files')
      .setText('Add timestamp: invoice.pdf → invoice_2025-10-04T14-30.pdf')
      .setSwitchControl(useOriginalSwitch);
    attachmentSection.addWidget(switchText);
    
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      var input = CardService.newTextInput()
        .setFieldName('newName' + i)
        .setValue(att.getName())
        .setHint('Leave blank to skip');
      attachmentSection.addWidget(input);
    }
    
    var preFilledCard = CardService.newCardBuilder()
      .setHeader(cardHeader)
      .addSection(folderSection)
      .addSection(attachmentSection)
      .setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
          .setText('Save to Drive')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('saveAttachments')
            .setParameters({ 'messageId': messageId, 'numAtts': numAtts.toString() }))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)));
    
    var navigation = CardService.newNavigation()
      .popToRoot()
      .pushCard(preFilledCard.build());
    
    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .build();
  } catch (error) {
    console.error('Error in selectFolder:', error);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popToRoot())
      .setNotification(CardService.newNotification().setText('Error selecting folder'))
      .build();
  }
}

function saveAttachments(e) {
  try {
    var accessToken = e.gmail.accessToken;
    var messageId = e.gmail.messageId;
    
    GmailApp.setCurrentMessageAccessToken(accessToken);
    
    var message = GmailApp.getMessageById(messageId);
    var attachments = message.getAttachments();
    var numAtts = attachments.length;
    
    var addTimestamp = e.formInputs.addTimestamp && e.formInputs.addTimestamp[0] === 'on';
    var folderId = e.formInputs.customFolderId ? e.formInputs.customFolderId[0] : null;
    
    if (!folderId) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('No folder selected'))
        .build();
    }
    
    var folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('Invalid folder'))
        .build();
    }
    
    var savedFiles = [];
    var savedCount = 0;
    var skippedCount = 0;
    
    for (var i = 0; i < numAtts; i++) {
      var newNameInput = e.formInputs['newName' + i] ? e.formInputs['newName' + i][0] : null;
      var finalName;
      
      if (!newNameInput || newNameInput.trim() === '') {
        skippedCount++;
        continue;
      }
      
      finalName = newNameInput.trim();
      
      if (addTimestamp) {
        var originalName = finalName;
        var dateSuffix = '_' + new Date().toISOString().substring(0, 16).replace(/:/g, '-');
        var lastDotIndex = originalName.lastIndexOf('.');
        if (lastDotIndex > 0) {
          finalName = originalName.substring(0, lastDotIndex) + dateSuffix + originalName.substring(lastDotIndex);
        } else {
          finalName = originalName + dateSuffix;
        }
      }
      
      if (finalName) {
        try {
          var newFile = folder.createFile(attachments[i].copyBlob());
          newFile.setName(finalName);
          savedFiles.push({ name: finalName, fileId: newFile.getId() });
          savedCount++;
        } catch (fileError) {
          console.error('Error saving file ' + i + ':', fileError);
          skippedCount++;
        }
      }
    }
    
    var successSection = CardService.newCardSection();
    
    if (savedCount > 0) {
      var statusWidget = CardService.newDecoratedText()
        .setTopLabel('✓ Success')
        .setText(savedCount + ' file' + (savedCount > 1 ? 's' : '') + ' saved to ' + folder.getName() + 
                (skippedCount > 0 ? ' (' + skippedCount + ' skipped)' : ''))
        .setWrapText(true);
      successSection.addWidget(statusWidget);
      successSection.addWidget(CardService.newDivider());
      
      for (var j = 0; j < Math.min(savedFiles.length, 5); j++) {
        var fileInfo = savedFiles[j];
        var openAction = CardService.newAction()
          .setFunctionName('openFileInDrive')
          .setParameters({ 'fileId': fileInfo.fileId });
        var fileWidget = CardService.newDecoratedText()
          .setText(fileInfo.name)
          .setButton(CardService.newTextButton()
            .setText('Open')
            .setOnClickAction(openAction));
        successSection.addWidget(fileWidget);
      }
      
      if (savedFiles.length > 5) {
        successSection.addWidget(CardService.newTextParagraph()
          .setText('<font color="#666666">+' + (savedFiles.length - 5) + ' more</font>'));
      }
    } else {
      var statusWidget = CardService.newDecoratedText()
        .setTopLabel('⚠ Notice')
        .setText('No files were saved')
        .setWrapText(true);
      successSection.addWidget(statusWidget);
      successSection.addWidget(CardService.newDivider());
    }
    
    var successCard = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Saved Successfully')
        .setSubtitle(savedCount > 0 ? savedCount + ' file' + (savedCount > 1 ? 's' : '') + ' saved' : 'No files saved')
        .setImageUrl('https://www.gstatic.com/images/icons/material/system/1x/check_circle_black_48dp.png'))
      .addSection(successSection)
      .setFixedFooter(CardService.newFixedFooter()
        .setPrimaryButton(CardService.newTextButton()
          .setText('Open Folder')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('openFolderInDrive')
            .setParameters({ 'folderId': folderId }))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .setSecondaryButton(CardService.newTextButton()
          .setText('Done')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('closeDialog'))));
    
    var navigation = CardService.newNavigation().pushCard(successCard.build());
    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .build();
      
  } catch (error) {
    console.error('Error in saveAttachments:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Save error'))
      .build();
  }
}

function openFolderInDrive(e) {
  var folderId = e.parameters.folderId;
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl('https://drive.google.com/drive/folders/' + folderId)
      .setOpenAs(CardService.OpenAs.FULL_SIZE)
      .setOnClose(CardService.OnClose.NOTHING))
    .build();
}

function openFileInDrive(e) {
  var fileId = e.parameters.fileId;
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl('https://drive.google.com/file/d/' + fileId + '/view')
      .setOpenAs(CardService.OpenAs.FULL_SIZE)
      .setOnClose(CardService.OnClose.NOTHING))
    .build();
}

function closeDialog(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popToRoot())
    .build();
}

function returnToHome(e) {
  try {
    var messageId = e.parameters.messageId;
    var numAtts = e.parameters.numAtts;
    
    var accessToken = null;
    if (e.gmail && e.gmail.accessToken) {
      accessToken = e.gmail.accessToken;
    }
    
    var mockE = {
      gmail: {
        messageId: messageId,
        accessToken: accessToken
      }
    };
    
    var mainCard = onGmailMessage(mockE);
    
    var navigation = CardService.newNavigation()
      .popToRoot()
      .pushCard(mainCard);
    
    return CardService.newActionResponseBuilder()
      .setNavigation(navigation)
      .build();
  } catch (error) {
    console.error('Error in returnToHome:', error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error returning to home. Please refresh.'))
      .build();
  }
}
