/* Copyright 2011 URL IS/IT
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you  may  not
 * use this file except in compliance with the License. You may obtain a copy of
 * the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless  required  by  applicable  law  or  agreed  to  in  writing,  software
 * distributed under the License is distributed on an  "AS  IS"  BASIS,  WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either  express  or  implied.  See  the
 * License for the specific language governing permissions and limitations under
 * the License.
 * 
 * Author Todd Url <toddurl@yahoo.com>
 * 
 * Url of the Development server - var url = "http://50.39.195.224:8888"
 */
var DEBUG = new Boolean(false);
var document = SpreadsheetApp.getActiveSpreadsheet();
var documentId = document.getId();
var documentUrl = document.getUrl();
var documentName = document.getName();
var url = "https://" + documentName + ".appspot.com";
var initializationUri = "/isInitialized";
var siteUpdateUri = "/" + documentId + "/updateSite";
var styleUpdateUri = "/" + documentId + "/updateStyle";
var landingUpdateUri = "/" + documentId + "/updateLanding";
var pageUpdateUri = "/" + documentId + "/updatePage";
var informationUpdateUri = "/" + documentId + "/updateItem";
var commitConfigurationUri = "/" + documentId + "/commitChange";
var rollbackConfigurationUri = "/" + documentId + "/rollbackConfiguration";
var siteConfigurationSheet = document.getSheetByName("SiteConfiguration");
var styleConfigurationSheet = document.getSheetByName("StyleConfiguration");
var landingConfigurationSheet = document.getSheetByName("LandingConfiguration");
var pageConfigurationSheet = document.getSheetByName("PageConfiguration");
var informationConfigurationSheet = document.getSheetByName("InformationConfiguration");
var menuEntries = [{name: "Initialize Configuration", functionName: "initialize"},
                   {name: "Update " + document.getName() + " configuration", functionName: "updateConfiguration"},
                   {name: "Display DocumentId", functionName: "displayConfigurationDocumentId"} ];

/*
 * onOpen()
 * 
 * Adds a main menu to the Apps Script Configuration Client.
 */
function onOpen() {
  document.addMenu("SitesWrapper", menuEntries);
}

/*
 * initialize()
 * 
 * Determines by HTTP response code if the datastore has been initialized with a DocumentId object and if not, sends
 * an mail to the Server Service Wrapper containing the document id of the spreadsheet managed by this Apps Script.
 * Once the DocumentId object exists, sheets in the spreadsheet are created and populated with a default configuration.
 */
function initialize () {
  var headers = {};
  var payload = {};
  var options = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:payload};
  var responseCode = UrlFetchApp.fetch(url + initializationUri, options).getResponseCode();
  if (responseCode == 204) {
    MailApp.sendEmail({
      to: "siteswrapper-gae-gwt@" + documentName + ".appspotmail.com",
      subject: documentName,
      body: documentId });
    while (responseCode != 202) {
      Utilities.sleep(500);
      responseCode = UrlFetchApp.fetch(url + initializationUri, options).getResponseCode();
    }
    initializeSite();
    initializeStyles();
    initializeLandings();
    initializeItems();
    initializeFirstPage();
    initializeSecondPage();
  } else if (responseCode == 202) {
    Browser.msgBox("Already Initialized");
  }
}

/*
 * Displays the Google Docs document id to the user. This id is unique and used as a key to enable the
 * configuration clients to update the datastore.
 */
function displayConfigurationDocumentId () {
  Browser.msgBox("The GoogleDocsConfigurationDocumentId for this webapp is " + documentId +
                 "The GoogleDocsConfigurationDocumentUrl for this webapp is " + documentUrl);
}

/*
 * Convenience method which sandboxes the external calls to UrlFetchApp in a try-catch block and
 * commits the new configuration to the datastore.
 */
function updateConfiguration() {
  try {
    updateSiteConfiguration();
    updateStyleConfiguration();
    updateLandingConfiguration();
    updatePageConfiguration();
    updateInformationConfiguration();
    commitConfigurationChanges();
  } catch(err) {
    if (err == "NO VALUE SPECIFIED FOR SITE NAME IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR Site Name IN SITECONFIGURATION WORKSHEET");
      return;
    } else if (err == "NO VALUE SPECIFIED FOR GOOGLE APP ENGINE APPLICATION IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR GOOGLE APP ENGINE APPLICATION IN SITECONFIGURATION WORKSHEET");
      return;
    } else if (err == "NO VALUE SPECIFIED FOR GOOGLE APP ENGINE VERSION IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR GOOGLE APP ENGINE VERSION IN SITECONFIGURATION WORKSHEET");
      return;
    } else if (err == "NO VALUE SPECIFIED FOR LOOK AND FEEL IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR LOOK AND FEEL IN SITECONFIGURATION WORKSHEET");
      return;
    } else if (err == "NO VALUE SPECIFIED FOR THEME IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR Theme IN SITECONFIGURATION WORKSHEET");
      return;
    } else if (err == "NO VALUE SPECIFIED FOR GOOGLE WEB FONTS URL IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR Google Web Fonts Url IN SITECONFIGURATION WORKSHEET");
      return;
    } else if (err == "NO VALUE SPECIFIED FOR FAVICON URL IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR Favicon Url IN SITECONFIGURATION WORKSHEET");
      return;
    } else if (err == "NO VALUE SPECIFIED FOR DEFAULT LANDING PAGE IN SITECONFIGURATION WORKSHEET") {
      Browser.msgBox("UPDATE ABORTED - NO VALUE SPECIFIED FOR Default Landing Page IN SiteConfiguration WORKSHEET");
      return;
    } else if (err == "DELETE OF CURRENT SITECONFIGURATION JDO IN APP ENGINE DATASTORE FAILED") {
      Browser.msgBox("UPDATE ABORTED - DELETE OF CURRENT SiteConfiguration JDO IN APP ENGINE DATASTORE FAILED");
      return;
    } else {
      Browser.msgBox(err);
      return;
    }
  }
}

/*
 * Updates the Site object in the datastore.
 */
function updateSiteConfiguration() {
  var configurationParameters = getColumnsData(siteConfigurationSheet, siteConfigurationSheet.getRange("B1:B" + siteConfigurationSheet.getLastRow()));
  var siteAttributes = "";
  if (typeof configurationParameters[0].siteName == "undefined") {
    throw "NO VALUE SPECIFIED FOR SITE NAME IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "siteName=" + escape(configurationParameters[0].siteName);
  }
  if (typeof configurationParameters[0].googleAppEngineApplication == "undefined") {
    throw "NO VALUE SPECIFIED FOR GOOGLE APP ENGINE APPLICATION IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "&googleAppEngineApplication=" + escape(configurationParameters[0].googleAppEngineApplication);
  }
  if (typeof configurationParameters[0].googleAppEngineVersion == "undefined") {
    throw "NO VALUE SPECIFIED FOR GOOGLE APP ENGINE VERSION IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "&googleAppEngineVersion=" + escape(configurationParameters[0].googleAppEngineVersion);
  }
  if (typeof configurationParameters[0].lookAndFeel == "undefined") {
    throw "NO VALUE SPECIFIED FOR LOOK AND FEEL IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "&lookAndFeel=" + escape(configurationParameters[0].lookAndFeel);
  }
  if (typeof configurationParameters[0].theme == "undefined") {
    throw "NO VALUE SPECIFIED FOR THEME IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "&theme=" + escape(configurationParameters[0].theme);
  }
  if (typeof configurationParameters[0].googleWebFontsUrl == "undefined") {
    throw "NO VALUE SPECIFIED FOR GOOGLE WEB FONTS URL IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "&googleWebFontsUrl=" + escape(configurationParameters[0].googleWebFontsUrl);
  }
  if (typeof configurationParameters[0].faviconUrl == "undefined") {
    throw "NO VALUE SPECIFIED FOR FAVICON URL IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "&faviconUrl=" + escape(configurationParameters[0].faviconUrl);
  }
  siteAttributes += "&appleTouchIconUrl=" + escape(configurationParameters[0].appleTouchIconUrl);
  if (typeof configurationParameters[0].defaultPage == "undefined") {
    throw "NO VALUE SPECIFIED FOR DEFAULT LANDING PAGE IN SITECONFIGURATION WORKSHEET";
  } else {
    siteAttributes += "&defaultPage=" + escape(configurationParameters[0].defaultPage);
  }
  if (escape(configurationParameters[0].revisionHistoryEnabled) == "Yes" || escape(configurationParameters[0].revisionHistoryEnabled) == "No") {
    siteAttributes += "&revisionHistoryEnabled=" + escape(configurationParameters[0].revisionHistoryEnabled);
  } else {
    throw "VALUE OF Yes OR No MUST BE SPECIFIED FOR Revision History Enabled IN SiteConfiguration WORKSHEET";
  }
  siteAttributes += "&logoImage=" + escape(configurationParameters[0].logoImage);
  siteAttributes += "&logoHtml=" + escape(configurationParameters[0].logoHtml);
  siteAttributes += "&displayLogoAs=" + escape(configurationParameters[0].displayLogoAs);
  if (typeof configurationParameters[0].siteFooter == "undefined") {
    throw "NO VALUE SPECIFIED FOR Site Footer IN SiteConfiguration WORKSHEET";
  } else {
    siteAttributes += "&siteFooter=" + escape(configurationParameters[0].siteFooter);
  }
  if (typeof configurationParameters[0].gwtRpcErrorMessage == "undefined") {
    throw "NO VALUE SPECIFIED FOR Gwt Rpc Error Message IN SiteConfiguration WORKSHEET";
  } else {
    siteAttributes += "&gwtRpcErrorMessage=" + escape(configurationParameters[0].gwtRpcErrorMessage);
  }
  var headers = {};
  headers.daoId = documentId;
  var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:siteAttributes};
  if (DEBUG == true) {
    Browser.msgBox("SiteConfiguration Attributes = " + siteAttributes);
  } else {
    if (UrlFetchApp.fetch(url + siteUpdateUri, advancedArguments).getContentText() != documentId) {
      throw "CREATION OF NEW SiteConfiguration OBJECT IN DATASTORE FAILED";
    }
  }
}

/*
 * Creates a new collection of Style objects in the datastore.
 */
function updateStyleConfiguration() {
  var maxColumns = styleConfigurationSheet.getLastColumn() + 1;
  var lookAndFeelTypes = getColumnsData(styleConfigurationSheet,
                                    styleConfigurationSheet.getRange("B1:" +
                                    styleConfigurationSheet.getLastColumn() +
                                    styleConfigurationSheet.getLastRow()));
  for (lookAndFeelType = 0; lookAndFeelType < lookAndFeelTypes.length; lookAndFeelType++) {
    var lookAndFeelTypeParameters = "";
    if (typeof lookAndFeelTypes[lookAndFeelType].lookAndFeel == "undefined") {
      throw "NO VALUE SPECIFIED FOR Look And Feel IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters = "lookAndFeel=" + escape(lookAndFeelTypes[lookAndFeelType].lookAndFeel);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].description == "undefined") {
      throw "NO VALUE SPECIFIED FOR Description IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&description=" + escape(lookAndFeelTypes[lookAndFeelType].description);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].primaryColor == "undefined") {
      throw "NO VALUE SPECIFIED FOR Primary Color IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&primaryColor=" + escape(lookAndFeelTypes[lookAndFeelType].primaryColor);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].primaryAccentColor == "undefined") {
      throw "NO VALUE SPECIFIED FOR Primary Accent Color IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&primaryAccentColor=" + escape(lookAndFeelTypes[lookAndFeelType].primaryAccentColor);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].secondaryAccentColor == "undefined") {
      throw "NO VALUE SPECIFIED FOR Secondary Accent Color IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&secondaryAccentColor=" + escape(lookAndFeelTypes[lookAndFeelType].secondaryAccentColor);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].tertiaryAccentColor == "undefined") {
      throw "NO VALUE SPECIFIED FOR Tertiary Accent Color IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&tertiaryAccentColor=" + escape(lookAndFeelTypes[lookAndFeelType].tertiaryAccentColor);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].mainMenuFontFamily == "undefined") {
      throw "NO VALUE SPECIFIED FOR Main Menu Font Family IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&mainMenuFontFamily=" + escape(lookAndFeelTypes[lookAndFeelType].mainMenuFontFamily);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].mainMenuFontSize == "undefined") {
      throw "NO VALUE SPECIFIED FOR Main Menu Font Size IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&mainMenuFontSize=" + escape(lookAndFeelTypes[lookAndFeelType].mainMenuFontSize);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].mainMenuSelectionFontColor == "undefined") {
      throw "NO VALUE SPECIFIED FOR Main Menu Selection Font Color IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&mainMenuSelectionFontColor=" + escape(lookAndFeelTypes[lookAndFeelType].mainMenuSelectionFontColor);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].mainMenuHoverFontColor == "undefined") {
      throw "NO VALUE SPECIFIED FOR Main Menu Hover Font Color IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&mainMenuHoverFontColor=" + escape(lookAndFeelTypes[lookAndFeelType].mainMenuHoverFontColor);
    }
    if (typeof lookAndFeelTypes[lookAndFeelType].mainMenuSelectedFontColor == "undefined") {
      throw "NO VALUE SPECIFIED FOR Main Menu Selected Font Color IN StyleConfiguration SHEET COLUMN " + lookAndFeelType;
    } else {
      lookAndFeelTypeParameters += "&mainMenuSelectedFontColor=" + escape(lookAndFeelTypes[lookAndFeelType].mainMenuSelectedFontColor);
    }
    if (DEBUG == true) {
      Browser.msgBox("InformationConfiguration Attributes = " + lookAndFeelTypeParameters);
    } else {
      var headers = {};
      headers.daoId = documentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:lookAndFeelTypeParameters};
      if (UrlFetchApp.fetch(url + styleUpdateUri, advancedArguments).getContentText() != documentId) {
        throw "CREATION OF NEW StyleConfiguration OBJECT IN APP ENGINE DATASTORE FAILED";
      } 
    }
  }
}

/*
 * Creates a new collection of Landing objects in the datastore.
 */
function updateLandingConfiguration() {
  var maxColumns = landingConfigurationSheet.getLastColumn() + 1;
  var landings = getColumnsData(landingConfigurationSheet,
                                    landingConfigurationSheet.getRange("B1:" +
                                    landingConfigurationSheet.getLastColumn() +
                                    landingConfigurationSheet.getLastRow()));
  for (landing = 0; landing < landings.length; landing++) {
    var landingParameters = "";
    if (typeof landings[landing].name == "undefined") {
      throw "NO VALUE SPECIFIED FOR Name IN LandingConfiguration SHEET COLUMN " + landing;
    } else {
      landingParameters = "name=" + escape(landings[landing].name);
    }
    if (typeof landings[landing].type == "undefined") {
      throw "NO VALUE SPECIFIED FOR Type IN LandingConfiguration SHEET COLUMN " + landing;
    } else {
      landingParameters += "&type=" + escape(landings[landing].type);
    }
    if (typeof landings[landing].description == "undefined") {
      throw "NO VALUE SPECIFIED FOR Description IN LandingConfiguration SHEET COLUMN " + landing;
    } else {
      landingParameters += "&description=" + escape(landings[landing].description);
    }
    if (typeof landings[landing].videoUrl == "undefined") {
      throw "NO VALUE SPECIFIED FOR Video Url IN LandingConfiguration SHEET COLUMN " + landing;
    } else {
      landingParameters += "&videoUrl=" + escape(landings[landing].videoUrl);
    }
    if (typeof landings[landing].imageUrl == "undefined") {
      throw "NO VALUE SPECIFIED FOR Image Url IN LandingConfiguration SHEET COLUMN " + landing;
    } else {
      landingParameters += "&imageUrl=" + escape(landings[landing].imageUrl);
    }
    if (typeof landings[landing].linkName == "undefined") {
      throw "NO VALUE SPECIFIED FOR Link Name IN LandingConfiguration SHEET COLUMN " + landing;
    } else {
      landingParameters += "&linkName=" + escape(landings[landing].linkName);
    }
    if (typeof landings[landing].linkUrl == "undefined") {
      throw "NO VALUE SPECIFIED FOR Link Url IN LandingConfiguration SHEET COLUMN " + landing;
    } else {
      landingParameters += "&linkUrl=" + escape(landings[landing].linkUrl);
    }
    landingParameters += "&specificationOne=" + escape(landings[landing].specificationOne);
    landingParameters += "&valueOne=" + escape(landings[landing].valueOne);
    landingParameters += "&specificationTwo=" + escape(landings[landing].specificationTwo);
    landingParameters += "&valueTwo=" + escape(landings[landing].valueTwo);
    landingParameters += "&specificationThree=" + escape(landings[landing].specificationThree);
    landingParameters += "&valueThree=" + escape(landings[landing].valueThree);
    landingParameters += "&specificationFour=" + escape(landings[landing].specificationFour);
    landingParameters += "&valueFour=" + escape(landings[landing].valueFour);
    landingParameters += "&specificationFive=" + escape(landings[landing].specificationFive);
    landingParameters += "&valueFive=" + escape(landings[landing].valueFive);
    landingParameters += "&specificationSix=" + escape(landings[landing].specificationSix);
    landingParameters += "&valueSix=" + escape(landings[landing].valueSix);
    landingParameters += "&specificationSeven=" + escape(landings[landing].specificationSeven);
    landingParameters += "&valueSeven=" + escape(landings[landing].valueSeven);
    landingParameters += "&specificationEight=" + escape(landings[landing].specificationEight);
    landingParameters += "&valueEight=" + escape(landings[landing].valueEight);
    landingParameters += "&specificationNine=" + escape(landings[landing].specificationNine);
    landingParameters += "&valueNine=" + escape(landings[landing].valueNine);
    landingParameters += "&specificationTen=" + escape(landings[landing].specificationTen);
    landingParameters += "&valueTen=" + escape(landings[landing].valueTen);
    if (DEBUG == true) {
      Browser.msgBox("LandingConfiguration Attributes = " + landingParameters);
    } else {
      var headers = {};
      headers.daoId = documentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:landingParameters};
      if (UrlFetchApp.fetch(url + landingUpdateUri, advancedArguments).getContentText() != documentId) {
        throw "CREATION OF NEW LandingConfiguration OBJECT IN APP ENGINE DATASTORE FOR PAGE FAILED";
      } 
    }
  }
}

/*
 * Creates a new collection of Page objects in the datastore.
 */
function updatePageConfiguration() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var pageSheets = new Array();
  var i = 0;
  for (var sheet = 0; sheet < sheets.length; sheet++) {
    if (sheets[sheet].getName() != "SiteConfiguration" &&
        sheets[sheet].getName() != "StyleConfiguration" &&
        sheets[sheet].getName() != "LandingConfiguration" &&
        sheets[sheet].getName() != "InformationConfiguration") {
      pageSheets[i++] = sheets[sheet];
    }
  }
  for (var sheet = 0; sheet < pageSheets.length; sheet++) {
    var maxColumns = pageSheets[sheet].getLastColumn() + 1;
    var configurationParameters = getColumnsData(pageSheets[sheet], pageSheets[sheet].getRange("B1:" + maxColumns + pageSheets[sheet].getLastRow()));
    var pageAttributes = "";
    if (typeof configurationParameters[0].pageName == "undefined") {
      throw "NO VALUE SPECIFIED FOR Page Name IN CONFIGURATION WORKSHEET " + pageSheets[sheet].pageName();
    } else {
      pageAttributes += "pageName=" + escape(configurationParameters[0].pageName);
    }
    if (escape(configurationParameters[0].showPageTitle) == "Yes" || escape(configurationParameters[0].showPageTitle) == "No") {
      pageAttributes += "&showPageTitle=" + escape(configurationParameters[0].showPageTitle);
    } else {
      throw "VALUE OF Yes OR No MUST BE SPECIFIED FOR Show Page Title IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    }
    pageAttributes += "&logoImage=" + escape(configurationParameters[0].logoImage);
    pageAttributes += "&logoHtml=" + escape(configurationParameters[0].logoHtml);
    if (escape(configurationParameters[0].displayLogoAs) == "Image" || escape(configurationParameters[0].displayLogoAs) == "Html") {
      pageAttributes += "&displayLogoAs=" + escape(configurationParameters[0].displayLogoAs);
    } else {
      throw "VALUE OF Image OR Html MUST BE SPECIFIED FOR Display Logo As IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    }
    if (escape(configurationParameters[0].languageTranslationEnabled) == "Yes" || escape(configurationParameters[0].languageTranslationEnabled) == "No") {
      pageAttributes += "&languageTranslationEnabled=" + escape(configurationParameters[0].languageTranslationEnabled);
    } else {
      throw "VALUE OF Yes OR No MUST BE SPECIFIED FOR Language Selection Enabled IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].translatedLanguages == "undefined")
          break;
        else
          pageAttributes += "&translatedLanguages=" + escape(configurationParameters[parameter].translatedLanguages);
      } catch (err) {
        break;
      }
    }
    if (configurationParameters[0].customSearchEnabled == "Yes" || configurationParameters[0].customSearchEnabled == "No") {
      pageAttributes += "&customSearchEnabled=" + escape(configurationParameters[0].customSearchEnabled);
    } else {
      throw "VALUE OF Yes OR No MUST BE SPECIFIED FOR Custom Search Enabled IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].customSearchUrls == "undefined")
          break;
        else
          pageAttributes += "&customSearchUrls=" + escape(configurationParameters[parameter].customSearchUrls);
      } catch (err) {
        break;
      }
    }
    if (configurationParameters[0].mainMenuType == "Link" || configurationParameters[0].mainMenuType == "Button") {
      pageAttributes += "&mainMenuType=" + escape(configurationParameters[0].mainMenuType);
    } else {
      throw "VALUE OF Link OR Button MUST BE SPECIFIED FOR Main Menu Type IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    }
    if (configurationParameters[0].mainMenuDirection == "Horizontal" || configurationParameters[0].mainMenuDirection == "Vertical" || configurationParameters[0].mainMenuDirection == "Both") {
      pageAttributes += "&mainMenuDirection=" + escape(configurationParameters[0].mainMenuDirection);
    } else {
      throw "VALUE OF Horizontal, Vertical OR Both MUST BE SPECIFIED FOR Main Menu Direction IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    }
    pageAttributes += "&mainMenuSelectionHtml=" + escape(configurationParameters[0].mainMenuSelectionHtml);
    pageAttributes += "&mainMenuSelectedHtml=" + escape(configurationParameters[0].mainMenuSelectedHtml);
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].backgroundImageUrls == "undefined")
          break;
        else
          pageAttributes += "&backgroundImageUrls=" + escape(configurationParameters[parameter].backgroundImageUrls);
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].backgroundImageDurationSeconds == "undefined")
          break;
        else
          pageAttributes += "&backgroundImageDurationSeconds=" + escape(configurationParameters[parameter].backgroundImageDurationSeconds);
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].contentMenuItemName == "undefined") {
          pageAttributes += "&contentMenuItemName=" + escape(configurationParameters[parameter].contentMenuItemName);
          break;
        } else {
          pageAttributes += "&contentMenuItemName=" + escape(configurationParameters[parameter].contentMenuItemName);
        }
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].contentMenuItemLink == "undefined") {
          pageAttributes += "&contentMenuItemLink=" + escape(configurationParameters[parameter].contentMenuItemLink);
          break;
        } else {
          pageAttributes += "&contentMenuItemLink=" + escape(configurationParameters[parameter].contentMenuItemLink);
        }
      } catch (err) {
        break;
      }
    }
    if (typeof configurationParameters[0].contentMenuStyleSheet != "undefined") {
      pageAttributes += "&contentMenuStyleSheet=" + escape(configurationParameters[0].contentMenuStyleSheet);
    } else {
      pageAttributes += "&contentMenuStyleSheet=" + escape(configurationParameters[0].contentMenuStyleSheet);
    }
    if (configurationParameters[0].contentLayout == "One column simple"    ||
        configurationParameters[0].contentLayout == "Two column simple"    ||
        configurationParameters[0].contentLayout == "Three column simple"  ||
        configurationParameters[0].contentLayout == "One column"           ||
        configurationParameters[0].contentLayout == "Two column"           ||
        configurationParameters[0].contentLayout == "Three column"         ||
        configurationParameters[0].contentLayout == "Left sidebar"         ||
        configurationParameters[0].contentLayout == "Right sidebar"        ||
        configurationParameters[0].contentLayout == "Left and right sidebars") {
      pageAttributes += "&contentLayout=" + escape(configurationParameters[0].contentLayout);
    } else {
      throw "VALUE OF One column simple,Two column simple,Three column simple,One column,Two column,Three column,Left sidebar,Right sidebar OR Left and right sidebar MUST BE SPECIFIED FOR Content Layout IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    }
    pageAttributes += "&contentHeader=" + escape(configurationParameters[0].contentHeader);
    pageAttributes += "&contentColumnOne=" + escape(configurationParameters[0].contentColumnOne);
    pageAttributes += "&contentColumnTwo=" + escape(configurationParameters[0].contentColumnTwo);
    pageAttributes += "&contentColumnThree=" + escape(configurationParameters[0].contentColumnThree);
    pageAttributes += "&contentLeftSidebar=" + escape(configurationParameters[0].contentLeftSidebar);
    pageAttributes += "&contentRightSidebar=" + escape(configurationParameters[0].contentRightSidebar);
    pageAttributes += "&contentFooter=" + escape(configurationParameters[0].contentFooter);
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messageHeaderText == "undefined") {
          var numMessages = parameter;
          break;
        } else {
          pageAttributes += "&messageHeaderText=" + escape(configurationParameters[parameter].messageHeaderText);
        }
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messageBodyText == "undefined")
          break;
        else
          pageAttributes += "&messageBodyText=" + escape(configurationParameters[parameter].messageBodyText);
      } catch (err) {
        break;
      }
    }
    
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messageInformationItem == "undefined")
          break;
        else
          pageAttributes += "&messageInformationItem=" + escape(configurationParameters[parameter].messageInformationItem);
      } catch (err) {
        break;
      }
    }
    
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messageHtmlColorCode == "undefined")
          break;
        else
          pageAttributes += "&messageHtmlColorCode=" + escape(configurationParameters[parameter].messageHtmlColorCode);
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messageWidthPercentOfPage == "undefined")
          break;
        else
          pageAttributes += "&messageWidthPercentOfPage=" + escape(configurationParameters[parameter].messageWidthPercentOfPage);
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messagePercentOfPageFromLeft == "undefined")
          break;
        else
          pageAttributes += "&messagePercentOfPageFromLeft=" + escape(configurationParameters[parameter].messagePercentOfPageFromLeft);
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messagePercentOfPageFromTop == "undefined")
          break;
        else
          pageAttributes += "&messagePercentOfPageFromTop=" + escape(configurationParameters[parameter].messagePercentOfPageFromTop);
      } catch (err) {
        break;
      }
    }
    for (var parameter = 0; parameter < maxColumns; parameter++) {
      try {
        if (typeof configurationParameters[parameter].messageDurationSeconds == "undefined")
          break;
        else
          pageAttributes += "&messageDurationSeconds=" + escape(configurationParameters[parameter].messageDurationSeconds);
      } catch (err) {
        break;
      }
    }
    if (escape(configurationParameters[0].contentLayout) == "undefined") {
      throw "NO VALUE SPECIFIED FOR Content Layout IN PageConfiguration WORKSHEET " + pageSheets[sheet].getName();
    } else {
      pageAttributes += "&contentLayout=" + escape(configurationParameters[0].contentLayout);
    }
    if (DEBUG == true) {
      Browser.msgBox("PageConfiguration Attributes Page=" + pageSheets[sheet].getName() + " and " + pageAttributes);
    } else {
      var headers = {};
      headers.daoId = documentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:pageAttributes};
      if (UrlFetchApp.fetch(url + pageUpdateUri, advancedArguments).getContentText() != documentId) {
        throw "CREATION OF NEW PageConfiguration OBJECT IN APP ENGINE DATASTORE FOR PAGE " + pageSheets[sheet].getName() + " FAILED";
      }
    }
  }
}

/*
 * Creates a new collection of Items in the datastore.
 */
function updateInformationConfiguration() {
  var maxColumns = informationConfigurationSheet.getLastColumn() + 1;
  var informationItems = getColumnsData(informationConfigurationSheet,
                                    informationConfigurationSheet.getRange("B1:" +
                                    informationConfigurationSheet.getLastColumn() +
                                    informationConfigurationSheet.getLastRow()));
  for (informationItem = 0; informationItem < informationItems.length; informationItem++) {
    var informationItemParameters = "";
    if (typeof informationItems[informationItem].name == "undefined") {
      throw "NO VALUE SPECIFIED FOR Name IN InformationConfiguration SHEET COLUMN " + informationItem;
    } else {
      informationItemParameters = "name=" + escape(informationItems[informationItem].name);
    }
    if (typeof informationItems[informationItem].type == "undefined") {
      throw "NO VALUE SPECIFIED FOR Type IN InformationConfiguration SHEET COLUMN " + informationItem;
    } else {
      informationItemParameters += "&type=" + escape(informationItems[informationItem].type);
    }
    if (typeof informationItems[informationItem].description == "undefined") {
      throw "NO VALUE SPECIFIED FOR Description IN InformationConfiguration SHEET COLUMN " + informationItem;
    } else {
      informationItemParameters += "&description=" + escape(informationItems[informationItem].description);
    }
    if (typeof informationItems[informationItem].videoUrl == "undefined") {
      throw "NO VALUE SPECIFIED FOR Video Url IN InformationConfiguration SHEET COLUMN " + informationItem;
    } else {
      informationItemParameters += "&videoUrl=" + escape(informationItems[informationItem].videoUrl);
    }
    if (typeof informationItems[informationItem].imageUrl == "undefined") {
      throw "NO VALUE SPECIFIED FOR Image Url IN InformationConfiguration SHEET COLUMN " + informationItem;
    } else {
      informationItemParameters += "&imageUrl=" + escape(informationItems[informationItem].imageUrl);
    }
    if (typeof informationItems[informationItem].linkName == "undefined") {
      throw "NO VALUE SPECIFIED FOR Link Name IN InformationConfiguration SHEET COLUMN " + informationItem;
    } else {
      informationItemParameters += "&linkName=" + escape(informationItems[informationItem].linkName);
    }
    if (typeof informationItems[informationItem].linkUrl == "undefined") {
      throw "NO VALUE SPECIFIED FOR Link Url IN InformationConfiguration SHEET COLUMN " + informationItem;
    } else {
      informationItemParameters += "&linkUrl=" + escape(informationItems[informationItem].linkUrl);
    }
    informationItemParameters += "&specificationOne=" + escape(informationItems[informationItem].specificationOne);
    informationItemParameters += "&valueOne=" + escape(informationItems[informationItem].valueOne);
    informationItemParameters += "&specificationTwo=" + escape(informationItems[informationItem].specificationTwo);
    informationItemParameters += "&valueTwo=" + escape(informationItems[informationItem].valueTwo);
    informationItemParameters += "&specificationThree=" + escape(informationItems[informationItem].specificationThree);
    informationItemParameters += "&valueThree=" + escape(informationItems[informationItem].valueThree);
    informationItemParameters += "&specificationFour=" + escape(informationItems[informationItem].specificationFour);
    informationItemParameters += "&valueFour=" + escape(informationItems[informationItem].valueFour);
    informationItemParameters += "&specificationFive=" + escape(informationItems[informationItem].specificationFive);
    informationItemParameters += "&valueFive=" + escape(informationItems[informationItem].valueFive);
    informationItemParameters += "&specificationSix=" + escape(informationItems[informationItem].specificationSix);
    informationItemParameters += "&valueSix=" + escape(informationItems[informationItem].valueSix);
    informationItemParameters += "&specificationSeven=" + escape(informationItems[informationItem].specificationSeven);
    informationItemParameters += "&valueSeven=" + escape(informationItems[informationItem].valueSeven);
    informationItemParameters += "&specificationEight=" + escape(informationItems[informationItem].specificationEight);
    informationItemParameters += "&valueEight=" + escape(informationItems[informationItem].valueEight);
    informationItemParameters += "&specificationNine=" + escape(informationItems[informationItem].specificationNine);
    informationItemParameters += "&valueNine=" + escape(informationItems[informationItem].valueNine);
    informationItemParameters += "&specificationTen=" + escape(informationItems[informationItem].specificationTen);
    informationItemParameters += "&valueTen=" + escape(informationItems[informationItem].valueTen);
    if (DEBUG == true) {
      Browser.msgBox("InformationConfiguration Attributes = " + informationItemParameters);
    } else {
      var headers = {};
      headers.daoId = documentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:informationItemParameters};
      if (UrlFetchApp.fetch(url + informationUpdateUri, advancedArguments).getContentText() != documentId) {
        throw "CREATION OF NEW InformationConfiguration OBJECT IN APP ENGINE DATASTORE FOR PAGE FAILED";
      } 
    }
  }
}

/*
 * Persists the newly created object to the datastore and alerts the user of completion.
 */
function commitConfigurationChanges() {
  var siteAttributes = "";
  var headers = {};
  headers.gDocsId = documentId;
  var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:siteAttributes};
  if (DEBUG != true) {
    if (UrlFetchApp.fetch(url + commitConfigurationUri, advancedArguments).getContentText() != documentId) {
      throw "COMMIT OF NEW CONFIGURATION IN APP ENGINE DATASTORE FOR PAGE FAILED";
    } else {
      Browser.msgBox("Update of " + document.getName() + " configuration successfull");
    }
  }
}

/*
 * initializeSite
 * 
 * Renames the initial sheet contained in a new spreadsheet to SiteConfiguration and populates
 * it with a default configuration.
 */
function initializeSite() {
  //var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  document.renameActiveSheet("SiteConfiguration");
  var sheet = document.getActiveSheet();
  sheet.appendRow(["Site Name", "My Site"]);
  sheet.appendRow(["Google App Engine Application", "towingenterpriseexecutive"]);
  sheet.appendRow(["Google App Engine Version", "1"]);
  sheet.appendRow(["Look And Feel", "Ghost"]);
  sheet.appendRow(["Theme", "Charcoal"]);
  sheet.appendRow(["Google Web Fonts Url", "http://fonts.googleapis.com/css?family=Aldrich|Raleway:100|Open+Sans:300,400"]);
  sheet.appendRow(["Favicon Url", "http://ghostgames.com/favicon.ico"]);
  sheet.appendRow(["Apple Touch Icon Url", "http://ssl.gstatic.com/sites/p/fff931/system/app/images/apple-touch-icon.png"]);
  sheet.appendRow(["Default Page", "About"]);
  sheet.appendRow(["Information Item Display Style", "Bottom"]);
  sheet.appendRow(["Revision History Enabled", "No"]);
  sheet.appendRow(["Logo Image", "http://googledrive.com/host/0B1wQZ0ttBuUaZVpyNkdKYnRobnc/Logo.png"]);
  sheet.appendRow(["Logo Html", "<h1><span style='font-family:Aldrich,arial,sans-serif;font-style:italic;font-weight:normal'><font color=#ffffff>The</font><font color=#00ff00>Green</font><font color=#ffffff>URL</font></span><sup><font color=#ffffff size=2>&reg;</font></sup></h1>"]);
  sheet.appendRow(["Display Logo As", "Html"]);
  sheet.appendRow(["Site Footer", "<p>My Site Footer</p>"]);
  sheet.appendRow(["Gwt Rpc Error Message", "Network Error - Check your network connection"]);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);
}

/*
 * initializeStyles
 * 
 * Creates a new sheet named StyleConfiguration and populates it with a default configuration.
 */
function initializeStyles() {
  document.insertSheet('StyleConfiguration', 1);
  sheet = document.setActiveSheet(document.getSheets()[1]);
  sheet.appendRow(["Look And Feel", "URL IS/IT", "Koninklijke", "Ghost"]);
  sheet.appendRow(["Description", "Looks like URL IS/IT's home page urlisit.com", "Reminiscent of www.usa.lighting.philips.com", "Classy minimalist back and white theme with red highlites in the spirit of Ghosts in Gothenburg ghostgames.com"]);
  sheet.appendRow(["Primary Color", "#101010", "#ffffff", "#000000"]);
  sheet.appendRow(["Primary Accent Color", "#d6d6d6", "#228B22", "#ffffff"]);
  sheet.appendRow(["Secondary Accent Color", "#aaaaaa", "#4169E1", "#ff0000"]);
  sheet.appendRow(["Tertiary Accent Color", "#eeeeee", "#00ff00", "#a4a4a4"]);
  sheet.appendRow(["Main Menu Font Family", "Open+Sans", "Raleway", "Open+Sans"]);
  sheet.appendRow(["Main Menu Font Size", "14px", "14px", "13px"]);
  sheet.appendRow(["Main Menu Selection Font Color", "#ffffff", "#228B22", "#ffffff"]);
  sheet.appendRow(["Main Menu Hover Font Color", "#0000ff", "#0000ff", "#fffc00"]);
  sheet.appendRow(["Main Menu Selected Font Color", "#fffc00", "#00ff00", "#ff0000"]);
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 300);
}
  
/*
 * initializeLandings
 * 
 * Creates a new sheet named LandingConfiguration and populates it with a default configuration.
 */
function initializeLandings() {
  document.insertSheet('LandingConfiguration', 2);
  sheet = document.setActiveSheet(document.getSheets()[2]);
  sheet.appendRow(["Name", "URLeCycle", "URLX-15", "VZ-8 SkyUTV", "SchwimmUTV"]);
  sheet.appendRow(["Type", "Page", "Page", "Page", "Page"]);
  sheet.appendRow(["Description", "Sporting a 1.21 gigawatt cobalt60 RTG powered DAYMAK front wheel, the URLiCycle is the ultimate in power-assist electric bicycles. In keeping with DAYMAK's clean simple aesthetic look, there are no break, gear, throttle or controller cables visible on the bike as it's completely 802.11n fly by wireless. Whether you live to ride, or ride to live, the URLeCycle is guaranteed to leave you breathless.", "With a Reaction Motors XLR-99 engine delivering 60,000 pounds of thrust and an Inconel-X heat-resistant fuselage, the URLX-15 is easly capable of attaining it's operational altitude of 60 miles or a top speed of 4,500 miles per hour.", "The Piasecki VZ-8 Sky UTV features two tandem, three-blade ducted rotors, with the crew of two seated between the two rotors. Power is handled by a Chevy 350 LT1 small block V8 piston engine, driving the rotors by a central gearbox.", "The Schwimmwagen amphibious UTV, which resembles a small highly manueverable 4-wheel drive sports car, is at home on water as it is in ruff terrain. It features a 4-stroke 4-cylinder horizontally-opposed air-cooled 1,131 cc German motor, 5 speed transaxle with ZF self-locking differentials on both the front and rear axles. When crossing water the three bade propeller is lowered from the rear deck engine cover and folded back up when not in use."]);
  sheet.appendRow(["Video Url", "http://youtu.be/_Ld83b7PC6w", "http://youtu.be/Jdq_l-8PNPA", "http://youtu.be/4SERvwWALOM", "http://youtu.be/A3ArELSi_K4"]);
  sheet.appendRow(["Image Url", "https://lh6.googleusercontent.com/-ZNl9jqIj5Hg/TepZHNGTWsI/AAAAAAAAAFU/XaTsnlcygLY/5.png", "http://lh6.googleusercontent.com/-CRkMobJTCsY/TekPSN6Ir3I/AAAAAAAAAEo/cJYgdcZyka8/3.png", "http://lh5.googleusercontent.com/-qj5ulShqEOo/TbgIHiby9PI/AAAAAAAAABc/JwNpT2j8AeA/1.png", "http://lh5.googleusercontent.com/-aTv3UQdMlgU/Te00HsQJMHI/AAAAAAAAAHE/lauRwOluexY/1.png"]);
  sheet.appendRow(["Link Name", "CRV Sales", "CRV Sales USA", "CRV Sales LLC", "Ironman"]);
  sheet.appendRow(["Link Url", "http://crvsalesusa.appspot.com/unavailableItem?item=8&image=3", "http://crvsalesusa.appspot.com/unavailableItem?item=10&image=2", "http://crvsalesusa.appspot.com/unavailableItem?item=9&image=5", "http://crvsalesusa.appspot.com/unavailableItem?item=11&image=1"]);
  sheet.appendRow(["Specification One", "Price", "Price", "Price", "Price"]);
  sheet.appendRow(["Value One", "$1,210,000.00", "Please call for availability and pricing", "$99,999.99", "Call"]);
  sheet.appendRow(["Specification Two", "", "", "", ""]);
  sheet.appendRow(["Value Two", "", "", "", ""]);
  sheet.appendRow(["Specification Three", "", "", "", ""]);
  sheet.appendRow(["Value Three", "", "", "", ""]);
  sheet.appendRow(["Specification Four", "", "", "", ""]);
  sheet.appendRow(["Value Four", "", "", "", ""]);
  sheet.appendRow(["Specification Five", "", "", "", ""]);
  sheet.appendRow(["Value Five", "", "", "", ""]);
  sheet.appendRow(["Specification Six", "", "", "", ""]);
  sheet.appendRow(["Value Six", "", "", "", ""]);
  sheet.appendRow(["Specification Seven", "", "", "", ""]);
  sheet.appendRow(["Value Seven", "", "", "", ""]);
  sheet.appendRow(["Specification Eight", "", "", "", ""]);
  sheet.appendRow(["Value Eight", "", "", "", ""]);
  sheet.appendRow(["Specification Nine", "", "", "", ""]);
  sheet.appendRow(["Value Nine", "", "", "", ""]);
  sheet.appendRow(["Specification Ten", "", "", "", ""]);
  sheet.appendRow(["Value Ten", "", "", "", ""]);
  sheet.setColumnWidth(1, 175);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 250);
}
  
/*
 * initializeItems
 * 
 * Creates a new sheet named LandingConfiguration and populates it with a default configuration.
 */
function initializeItems() {
  document.insertSheet('InformationConfiguration', 3);
  sheet = document.setActiveSheet(document.getSheets()[3]);
  sheet.appendRow(["Name", "URLeCycle", "URLX-15", "VZ-8 SkyUTV", "SchwimmUTV"]);
  sheet.appendRow(["Type", "Page", "Page", "Page", "Page"]);
  sheet.appendRow(["Description", "Sporting a 1.21 gigawatt cobalt60 RTG powered DAYMAK front wheel, the URLiCycle is the ultimate in power-assist electric bicycles. In keeping with DAYMAK's clean simple aesthetic look, there are no break, gear, throttle or controller cables visible on the bike as it's completely 802.11n fly by wireless. Whether you live to ride, or ride to live, the URLeCycle is guaranteed to leave you breathless.", "With a Reaction Motors XLR-99 engine delivering 60,000 pounds of thrust and an Inconel-X heat-resistant fuselage, the URLX-15 is easly capable of attaining it's operational altitude of 60 miles or a top speed of 4,500 miles per hour.", "The Piasecki VZ-8 Sky UTV features two tandem, three-blade ducted rotors, with the crew of two seated between the two rotors. Power is handled by a Chevy 350 LT1 small block V8 piston engine, driving the rotors by a central gearbox.", "The Schwimmwagen amphibious UTV, which resembles a small highly manueverable 4-wheel drive sports car, is at home on water as it is in ruff terrain. It features a 4-stroke 4-cylinder horizontally-opposed air-cooled 1,131 cc German motor, 5 speed transaxle with ZF self-locking differentials on both the front and rear axles. When crossing water the three bade propeller is lowered from the rear deck engine cover and folded back up when not in use."]);
  sheet.appendRow(["Video Url", "http://youtu.be/_Ld83b7PC6w", "http://youtu.be/Jdq_l-8PNPA", "http://youtu.be/4SERvwWALOM", "http://youtu.be/A3ArELSi_K4"]);
  sheet.appendRow(["Image Url", "https://lh6.googleusercontent.com/-ZNl9jqIj5Hg/TepZHNGTWsI/AAAAAAAAAFU/XaTsnlcygLY/5.png", "http://lh6.googleusercontent.com/-CRkMobJTCsY/TekPSN6Ir3I/AAAAAAAAAEo/cJYgdcZyka8/3.png", "http://lh5.googleusercontent.com/-qj5ulShqEOo/TbgIHiby9PI/AAAAAAAAABc/JwNpT2j8AeA/1.png", "http://lh5.googleusercontent.com/-aTv3UQdMlgU/Te00HsQJMHI/AAAAAAAAAHE/lauRwOluexY/1.png"]);
  sheet.appendRow(["Link Name", "CRV Sales", "CRV Sales USA", "CRV Sales LLC", "Ironman"]);
  sheet.appendRow(["Link Url", "http://crvsalesusa.appspot.com/unavailableItem?item=8&image=3", "http://crvsalesusa.appspot.com/unavailableItem?item=10&image=2", "http://crvsalesusa.appspot.com/unavailableItem?item=9&image=5", "http://crvsalesusa.appspot.com/unavailableItem?item=11&image=1"]);
  sheet.appendRow(["Specification One", "Price", "Price", "Price", "Price"]);
  sheet.appendRow(["Value One", "$1,210,000.00", "Please call for availability and pricing", "$99,999.99", "Call"]);
  sheet.appendRow(["Specification Two", "", "", "", ""]);
  sheet.appendRow(["Value Two", "", "", "", ""]);
  sheet.appendRow(["Specification Three", "", "", "", ""]);
  sheet.appendRow(["Value Three", "", "", "", ""]);
  sheet.appendRow(["Specification Four", "", "", "", ""]);
  sheet.appendRow(["Value Four", "", "", "", ""]);
  sheet.appendRow(["Specification Five", "", "", "", ""]);
  sheet.appendRow(["Value Five", "", "", "", ""]);
  sheet.appendRow(["Specification Six", "", "", "", ""]);
  sheet.appendRow(["Value Six", "", "", "", ""]);
  sheet.appendRow(["Specification Seven", "", "", "", ""]);
  sheet.appendRow(["Value Seven", "", "", "", ""]);
  sheet.appendRow(["Specification Eight", "", "", "", ""]);
  sheet.appendRow(["Value Eight", "", "", "", ""]);
  sheet.appendRow(["Specification Nine", "", "", "", ""]);
  sheet.appendRow(["Value Nine", "", "", "", ""]);
  sheet.appendRow(["Specification Ten", "", "", "", ""]);
  sheet.appendRow(["Value Ten", "", "", "", ""]);
  sheet.setColumnWidth(1, 175);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 250);
}

/*
 * initializeFirstPage
 * 
 * Creates a new sheet with an arbitrary name (the name of a page in the site) and populates it
 * with a default configuration.
 */
function initializeFirstPage() {
  document.insertSheet('About', 4);
  sheet = document.setActiveSheet(document.getSheets()[4]);
  sheet.appendRow(["Page Name", "About"]);
  sheet.appendRow(["Show Page Title", "Yes"]);
  sheet.appendRow(["Logo Image", "https://c824ff113391b7c600d1069f19350d6607b580e1.googledrive.com/host/0BzPelJUA_7zUT3ZfQVdNcmwzbDg/SitesWrapperLogoLarge300x39.png", "", "", ""]);
  sheet.appendRow(["Logo Html", "<h1><span style='font-family:Aldrich,arial,sans-serif;font-style:italic;font-weight:normal'><font color=#ffffff>The</font><font color=#ff0000>Red</font><font color=#ffffff>URL</font></span><sup><font color=#ffffff size=2>&reg;</font></sup></h1>"]);
  sheet.appendRow(["Display Logo As", "Image"]);
  sheet.appendRow(["Language Translation Enabled", "No"]);
  sheet.appendRow(["Translated Languages", "en", "es"]);
  sheet.appendRow(["Custom Search Enabled", "Yes"]);
  sheet.appendRow(["Custom Search Urls", "<form id=sites-searchbox-form action=/site/poultryledlighting/system/app/pages/search><input type=hidden id=sites-searchbox-scope name=scope value=search-site /><input type=text id=jot-ui-searchInput name=q size=20 value= aria-label=\"Search this site\" autocomplete=off /><div id=sites-searchbox-button-set class=goog-inline-block><div role=button id=sites-searchbox-search-button class=\"goog-inline-block jfk-button jfk-button-standard\" tabindex=0>Search</div></div></form>"]);
  sheet.appendRow(["Main Menu Type", "Link"]);
  sheet.appendRow(["Main Menu Direction", "Horizontal"]);
  sheet.appendRow(["Main Menu Selection Html"]);
  sheet.appendRow(["Main Menu Selected Html"]);
  sheet.appendRow(["Background Image Urls", "/images/BackgroundImage02.jpg", "/images/GreaterWidthRatioScaleUp.png", "http://www.spektyr.com/PrintImages/Cerulean%20Cross%203%20Large.jpg"]);
  sheet.appendRow(["Background Image Duration Seconds", "5", "5", "5"]);
  sheet.appendRow(["Content Menu Item Name", "Solid State Lighting", "", "", ""]);
  sheet.appendRow(["Content Menu Item Link", "http://sites.google.com/site/solidstatelamps/", "", "", ""]);
  sheet.appendRow(["Content Menu Style Sheet", "", "", "", ""]);
  sheet.appendRow(["Content Layout", "Left sidebar", "", "", ""]);
  sheet.appendRow(["Content Header", "<span style=\"font-size:24px\">Introducing the UR<font color=\"#00ff00\">LeD</font>&trade; 100,000 hour solid state A19 lamp from <span style=\"font-family:Aldrich,arial,sans-serif;font-style:italic;font-weight:normal\"><font color=\"#ffffff\">The</font><font color=\"#00ff00\">Green</font><font color=\"#ffffff\">URL</font></span><sup><font color=\"#ffffff\" size=\"2\">&reg;</font></sup></span>", "", "", ""]);
  sheet.appendRow(["Content Column One", "UR<font color=\"#00ff00\">LeD</font>&trade; 6 watt lamps are the first 60 watt incandescent or 12 watt compact fluorescent general purpose replacement bulb to offer a lifetime warranty. Relamping with a&nbsp;UR<font color=\"#00ff00\">LeD</font>&trade; solid state lamp can save as much as $10 a year in electricity and may even eliminate the need to buy a bulb again.  <div><br> </div> <div>With most lamps lifespan depends on many factors, but in the case of LED lighting, it's all about thermal management. Like all semiconductors, solid state light emitting diods are degraded or damaged as the result of operating at high temperatures, not operating for long periods of time or being switched on and off. That's why <font color=\"#00ff00\">U</font>niversally <font color=\"#00ff00\">R</font>ecyclable <font color=\"#00ff00\">L</font>ighting&trade; puts 2.2 onces of aluminum heat sink at the core of each lamp, in order to maintain a low junction temperature and thus ensure high performance, high efficiency and long life out of each of the 77 Epistar Superbright SMD 3528 LEDs.</div> <div><br> </div> <div>A UR<font color=\"#00ff00\">LeD</font>&trade; lamp produces as much light as a conventional 60 watt light bulb yet only uses 6 watts and doesn't contain harmfull mercury or produce ultraviolet radiation. Plus, if your lamp ever burns out or diminishes in luminosity by more than 25% we'll replace it free of charge. That means that at todays energy prices a UR<font color=\"#00ff00\">LeD</font>&trade; lamp will typically pay for itself within two years, and since it doesn't need to be replaced, the savings won't stop there.</div> <div><br> </div> <div>Switching from traditional light bulbs to solid state lighting may seem like a burdon at first, but it doesn't have to be. After all, a light bulb's closest relative is the vacum tube, and if history has anything to teach us it's that \"in general\" we're better off with solid state TV's, flat screen monitors and digital cameras than we were with vacum tubes. The same is true with solid state lighting. Burned out bulbs, dimly lit homes and remembering to turn off the lights can literally be a thing of the past. After all, at 6 watts a UR<font color=\"#00ff00\">LeD</font>&trade; lamp uses less electricity than a typical night light.</div> <div><br> </div> <div>We want to help you step into the future of lighting with Solid State Lamps by making it as inexpensive, easy and risk free as possible. That's why we've selected UR<font color=\"#00ff00\">LeD</font>&trade; lamps and are offering them to you in the following packages, each of which comes with our unprecedented warranty and free shipping as well as increased savings over the previous package.</div> <div> <hr> <br> </div> <table border=\"1\" style=\"width:100%;border-collapse:collapse\"> <tbody> <tr> <td style=\"width:33.33%;background-color:#606060;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">One UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; lamp only $20</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Start saving today</font></font></div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/1_Lamp.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/1_Lamp.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">One UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; 100,000 hour A19 solid state LED bulb with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://www.sandbox.paypal.com/cgi-bin/webscr?cmd=_s-xclick&amp;hosted_button_id=EXMWD68889Z46\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#707070;text-align:center\"> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; lamps only $80</font> </font> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Save $20</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/5_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/5_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; 100,000 hour A19 solid state LED light bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#808080;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Ten UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; lamps only $140</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Save $60</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/10_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/10_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Ten UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; 100,000 hour A19 solid state LED light bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> </tr> <tr> <td style=\"width:33.33%;background-color:#909090;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Fifteen UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; lamps only $200</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Save $100</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/15_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/15_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Fifteen UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; 100,000 hour A19 solid state LED light bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#a0a0a0;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; lamps only $260</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Save $140</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/20_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/20_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty UR</font><font color=\"#00ff00\">LeD</font> <font color=\"#000000\">&trade; 100,000 hour A19 solid state LED bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#b0b0b0;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; lamps only $320</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Save $180</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/25_Lamps3.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/25_Lamps3.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; 100,000 hour A19 solid state LED bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> </tr> </tbody> </table>", "", "", ""]);
  sheet.appendRow(["Content Column Two", "", "", "", ""]);
  sheet.appendRow(["Content Column Three", "", "", "", ""]);
  sheet.appendRow(["Content Left Sidebar", "<div style=\"display:block;text-align:left\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/technology/Logo6.png\" style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> </div> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-weight:bold\"><span style=\"font-size:medium\"><span style=\"font-family:arial,sans-serif\"><font><font color=\"#134f5c\"><br> </font></font></span></span></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-weight:bold\"><span style=\"font-size:medium\"><span style=\"font-family:arial,sans-serif\"><font><font color=\"#134f5c\">Store Hours</font></font></span></span></span></p> <hr> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\">Monday-Friday&nbsp;</font></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\"><span style=\"font-size:x-small\">9:00AM - 5:00PM</span></font></span> </p> <font face=\"arial, sans-serif\"> <hr> </font> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\">Saturday&nbsp;</font></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><span style=\"font-size:x-small\"><font color=\"#444444\">10:00AM - 4:00PM</font></span></span></p> <font face=\"arial, sans-serif\"> <hr> </font> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\">Sunday&nbsp;</font></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><span style=\"font-size:x-small\"><font color=\"#444444\">12:00PM - 6:00PM</font></span></span></p> <hr> <div style=\"display:block;text-align:left\"> <div style=\"display:block;text-align:left\"><img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/PayPal2.png\" style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"></div> </div> <div style=\"text-align:left\"><br> </div> ", "", "", ""]);
  sheet.appendRow(["Content Right Sidebar", "", "", "", ""]);
  sheet.appendRow(["Content Footer", "<span style=\"font-size:24px\">Introducing the UR<font color=\"#00ff00\">LeD</font>&trade; 100,000 hour solid state A19 lamp from <span style=\"font-family:Aldrich,arial,sans-serif;font-style:italic;font-weight:normal\"><font color=\"#ffffff\">The</font><font color=\"#00ff00\">Green</font><font color=\"#ffffff\">URL</font></span><sup><font color=\"#ffffff\" size=\"2\">&reg;</font></sup></span>", "", "", ""]);
  sheet.appendRow(["Message Header Text", "<font color=#ffc000>The easiest way to create a beautiful enterprise class web application</font>", "Five URLeD™ lamp only $80 save $20", "", ""]);
  sheet.appendRow(["Message Body Text", "Whether you need simple pages, striking galleries, a professional blog, or an online store, it's all possible with SitesWrapper. Everything is mobile-ready and search engine optimized. Best of all, it's free!", "", ""]);
  sheet.appendRow(["Message Information Item", "URLeCycle", "none", "", ""]);
  sheet.appendRow(["Message Html Color Code", "#cccccc", "#cccccc", "", ""]);
  sheet.appendRow(["Message Width Percent Of Page", "0.15", "0.05", "", ""]);
  sheet.appendRow(["Message Percent Of Page From Left", "0.05", "0.2", "", ""]);
  sheet.appendRow(["Message Percent Of Page From Top", "0.25", "0.5", "", ""]);
  sheet.appendRow(["Message Duration Seconds", "5", "0.05", "", ""]);
  sheet.setColumnWidth(1, 225);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
}

/*
 * initializeSecondPage
 * 
 * Creates a new sheet with an arbitrary name (a page in the site) and populates it with a default configuration.
 */
function initializeSecondPage() {
  document.insertSheet('Themes', 5);
  sheet = document.setActiveSheet(document.getSheets()[5]);
  sheet.appendRow(["Page Name", "Themes"]);
  sheet.appendRow(["Show Page Title", "Yes"]);
  sheet.appendRow(["Logo Image", "https://c824ff113391b7c600d1069f19350d6607b580e1.googledrive.com/host/0BzPelJUA_7zUT3ZfQVdNcmwzbDg/SitesWrapperLogoLarge300x39.png", "", "", ""]);
  sheet.appendRow(["Logo Html", "<h1><span style='font-family:Aldrich,arial,sans-serif;font-style:italic;font-weight:normal'><font color=#ffffff>The</font><font color=#ff0000>Red</font><font color=#ffffff>URL</font></span><sup><font color=#ffffff size=2>&reg;</font></sup></h1>"]);
  sheet.appendRow(["Display Logo As", "Image"]);
  sheet.appendRow(["Language Translation Enabled", "No"]);
  sheet.appendRow(["Translated Languages", "en", "es"]);
  sheet.appendRow(["Custom Search Enabled", "Yes"]);
  sheet.appendRow(["Custom Search Urls", "<form id=sites-searchbox-form action=/site/poultryledlighting/system/app/pages/search><input type=hidden id=sites-searchbox-scope name=scope value=search-site /><input type=text id=jot-ui-searchInput name=q size=20 value= aria-label=\"Search this site\" autocomplete=off /><div id=sites-searchbox-button-set class=goog-inline-block><div role=button id=sites-searchbox-search-button class=\"goog-inline-block jfk-button jfk-button-standard\" tabindex=0>Search</div></div></form>"]);
  sheet.appendRow(["Main Menu Type", "Link"]);
  sheet.appendRow(["Main Menu Direction", "Horizontal"]);
  sheet.appendRow(["Main Menu Selection Html"]);
  sheet.appendRow(["Main Menu Selected Html"]);
  sheet.appendRow(["Background Image Urls", "/images/SitesWrapperAbout.jpg", "/images/GreaterWidthRatioScaleUp.png", "http://www.spektyr.com/PrintImages/Cerulean%20Cross%203%20Large.jpg"]);
  sheet.appendRow(["Background Image Duration Seconds", "5", "5", "5"]);
  sheet.appendRow(["Content Menu Item Name", "Solid State Lighting", "", "", ""]);
  sheet.appendRow(["Content Menu Item Link", "http://sites.google.com/site/solidstatelamps/", "", "", ""]);
  sheet.appendRow(["Content Menu Style Sheet", "", "", "", ""]);
  sheet.appendRow(["Content Layout", "Left sidebar", "", "", ""]);
  sheet.appendRow(["Content Header", "<span style=\"font-size:24px\">Introducing the UR<font color=\"#00ff00\">LeD</font>&trade; 100,000 hour solid state A19 lamp from <span style=\"font-family:Aldrich,arial,sans-serif;font-style:italic;font-weight:normal\"><font color=\"#ffffff\">The</font><font color=\"#00ff00\">Green</font><font color=\"#ffffff\">URL</font></span><sup><font color=\"#ffffff\" size=\"2\">&reg;</font></sup></span>", "", "", ""]);
  sheet.appendRow(["Content Column One", "UR<font color=\"#00ff00\">LeD</font>&trade; 6 watt lamps are the first 60 watt incandescent or 12 watt compact fluorescent general purpose replacement bulb to offer a lifetime warranty. Relamping with a&nbsp;UR<font color=\"#00ff00\">LeD</font>&trade; solid state lamp can save as much as $10 a year in electricity and may even eliminate the need to buy a bulb again.  <div><br> </div> <div>With most lamps lifespan depends on many factors, but in the case of LED lighting, it's all about thermal management. Like all semiconductors, solid state light emitting diods are degraded or damaged as the result of operating at high temperatures, not operating for long periods of time or being switched on and off. That's why <font color=\"#00ff00\">U</font>niversally <font color=\"#00ff00\">R</font>ecyclable <font color=\"#00ff00\">L</font>ighting&trade; puts 2.2 onces of aluminum heat sink at the core of each lamp, in order to maintain a low junction temperature and thus ensure high performance, high efficiency and long life out of each of the 77 Epistar Superbright SMD 3528 LEDs.</div> <div><br> </div> <div>A UR<font color=\"#00ff00\">LeD</font>&trade; lamp produces as much light as a conventional 60 watt light bulb yet only uses 6 watts and doesn't contain harmfull mercury or produce ultraviolet radiation. Plus, if your lamp ever burns out or diminishes in luminosity by more than 25% we'll replace it free of charge. That means that at todays energy prices a UR<font color=\"#00ff00\">LeD</font>&trade; lamp will typically pay for itself within two years, and since it doesn't need to be replaced, the savings won't stop there.</div> <div><br> </div> <div>Switching from traditional light bulbs to solid state lighting may seem like a burdon at first, but it doesn't have to be. After all, a light bulb's closest relative is the vacum tube, and if history has anything to teach us it's that \"in general\" we're better off with solid state TV's, flat screen monitors and digital cameras than we were with vacum tubes. The same is true with solid state lighting. Burned out bulbs, dimly lit homes and remembering to turn off the lights can literally be a thing of the past. After all, at 6 watts a UR<font color=\"#00ff00\">LeD</font>&trade; lamp uses less electricity than a typical night light.</div> <div><br> </div> <div>We want to help you step into the future of lighting with Solid State Lamps by making it as inexpensive, easy and risk free as possible. That's why we've selected UR<font color=\"#00ff00\">LeD</font>&trade; lamps and are offering them to you in the following packages, each of which comes with our unprecedented warranty and free shipping as well as increased savings over the previous package.</div> <div> <hr> <br> </div> <table border=\"1\" style=\"width:100%;border-collapse:collapse\"> <tbody> <tr> <td style=\"width:33.33%;background-color:#606060;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">One UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; lamp only $20</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Start saving today</font></font></div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/1_Lamp.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/1_Lamp.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">One UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; 100,000 hour A19 solid state LED bulb with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://www.sandbox.paypal.com/cgi-bin/webscr?cmd=_s-xclick&amp;hosted_button_id=EXMWD68889Z46\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#707070;text-align:center\"> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; lamps only $80</font> </font> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Save $20</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/5_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/5_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; 100,000 hour A19 solid state LED light bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#808080;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Ten UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; lamps only $140</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Save $60</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/10_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/10_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#ffffff\">Ten UR</font><font color=\"#00ff00\">LeD</font><font color=\"#ffffff\">&trade; 100,000 hour A19 solid state LED light bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> </tr> <tr> <td style=\"width:33.33%;background-color:#909090;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Fifteen UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; lamps only $200</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Save $100</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/15_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/15_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Fifteen UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; 100,000 hour A19 solid state LED light bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#a0a0a0;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; lamps only $260</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Save $140</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/20_Lamps.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/20_Lamps.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty UR</font><font color=\"#00ff00\">LeD</font> <font color=\"#000000\">&trade; 100,000 hour A19 solid state LED bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> <td style=\"width:33.33%;background-color:#b0b0b0;text-align:center\"> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; lamps only $320</font> </font> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Save $180</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"https://sites.google.com/site/solidstatelamps/home/25_Lamps3.png?attredirects=0\" imageanchor=\"1\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/25_Lamps3.png\"> </a> </div> <div> <font size=\"2\" style=\"font-style:normal\"> <font color=\"#000000\">Twenty five UR</font><font color=\"#00ff00\">LeD</font><font color=\"#000000\">&trade; 100,000 hour A19 solid state LED bulbs with lifetime warranty and free shipping in the</font>&nbsp; <font color=\"#ff0000\">U</font> <font color=\"#ffffff\">S</font> <font color=\"#0000ff\">A</font> </font> </div> <div style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> <a href=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\" imageanchor=\"1\"> <img border=\"0\" src=\"http://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif\"> </a> </div> </td> </tr> </tbody> </table>", "", "", ""]);
  sheet.appendRow(["Content Column Two", "", "", "", ""]);
  sheet.appendRow(["Content Column Three", "", "", "", ""]);
  sheet.appendRow(["Content Left Sidebar", "<div style=\"display:block;text-align:left\"> <img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/technology/Logo6.png\" style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"> </div> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-weight:bold\"><span style=\"font-size:medium\"><span style=\"font-family:arial,sans-serif\"><font><font color=\"#134f5c\"><br> </font></font></span></span></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-weight:bold\"><span style=\"font-size:medium\"><span style=\"font-family:arial,sans-serif\"><font><font color=\"#134f5c\">Store Hours</font></font></span></span></span></p> <hr> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\">Monday-Friday&nbsp;</font></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\"><span style=\"font-size:x-small\">9:00AM - 5:00PM</span></font></span> </p> <font face=\"arial, sans-serif\"> <hr> </font> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\">Saturday&nbsp;</font></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><span style=\"font-size:x-small\"><font color=\"#444444\">10:00AM - 4:00PM</font></span></span></p> <font face=\"arial, sans-serif\"> <hr> </font> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;font-size:14px;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><font color=\"#444444\">Sunday&nbsp;</font></span></p> <p style=\"text-align:center;margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;font-style:normal;font-variant:normal;font-weight:normal;line-height:normal\"><span style=\"font-family:arial,sans-serif\"><span style=\"font-size:x-small\"><font color=\"#444444\">12:00PM - 6:00PM</font></span></span></p> <hr> <div style=\"display:block;text-align:left\"> <div style=\"display:block;text-align:left\"><img border=\"0\" src=\"https://sites.google.com/site/solidstatelamps/home/PayPal2.png\" style=\"display:block;margin-right:auto;margin-left:auto;text-align:center\"></div> </div> <div style=\"text-align:left\"><br> </div> ", "", "", ""]);
  sheet.appendRow(["Content Right Sidebar", "", "", "", ""]);
  sheet.appendRow(["Content Footer", "<span style=\"font-size:24px\">Introducing the UR<font color=\"#00ff00\">LeD</font>&trade; 100,000 hour solid state A19 lamp from <span style=\"font-family:Aldrich,arial,sans-serif;font-style:italic;font-weight:normal\"><font color=\"#ffffff\">The</font><font color=\"#00ff00\">Green</font><font color=\"#ffffff\">URL</font></span><sup><font color=\"#ffffff\" size=\"2\">&reg;</font></sup></span>", "", "", ""]);
  sheet.appendRow(["Message Header Text", "<font color=#ffc000>SitesWrapper makes it simple to create a beautiful, professional web presence</font>", "Five URLeD™ lamp only $80 save $20", "", ""]);
  sheet.appendRow(["Message Body Text", "Promote your business, showcase your art, set up an online shop or just sharpen your Java programming skills. SitesWrapper is a website builder that has everything you need to build enterprise class web presence free. Browse our collection of beautiful website templates. You'll find loads of stunning designs, ready to be customized.", "", ""]);
  sheet.appendRow(["Message Information Item", "URLeCycle", "none", "", ""]);
  sheet.appendRow(["Message Html Color Code", "#cccccc", "#cccccc", "", ""]);
  sheet.appendRow(["Message Width Percent Of Page", "0.15", "0.05", "", ""]);
  sheet.appendRow(["Message Percent Of Page From Left", "0.66", "0.2", "", ""]);
  sheet.appendRow(["Message Percent Of Page From Top", "0.33", "0.5", "", ""]);
  sheet.appendRow(["Message Duration Seconds", "5", "0.05", "", ""]);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
}

/*
 * Google provided example functions below.
 */
// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}

// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range; 
// Returns an Array of objects.
function getColumnsData(sheet, range, rowHeadersColumnIndex) {
  rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
  var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
  var headers = normalizeHeaders(arrayTranspose(headersTmp)[0]);
  return getObjects(arrayTranspose(range.getValues()), headers);
}