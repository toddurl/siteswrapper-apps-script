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
 * the License. */

var DEBUG = new Boolean(false);
//var siteUrl = "http://50.39.195.224:8888"
var siteUrl = "https://urlisit.appspot.com"
var configurationDocument = SpreadsheetApp.getActiveSpreadsheet();
var configurationDocumentId = configurationDocument.getId();
var configurationDocumentUrl = configurationDocument.getUrl();
var siteUpdateUri = "/" + configurationDocumentId + "/updateSite";
var styleUpdateUri = "/" + configurationDocumentId + "/updateStyle";
var landingUpdateUri = "/" + configurationDocumentId + "/updateLanding";
var pageUpdateUri = "/" + configurationDocumentId + "/updatePage";
var informationUpdateUri = "/" + configurationDocumentId + "/updateItem";
var commitConfigurationUri = "/" + configurationDocumentId + "/commitChange";
var rollbackConfigurationUri = "/" + configurationDocumentId + "/rollbackConfiguration";
var siteConfigurationSheet = configurationDocument.getSheetByName("SiteConfiguration");
var styleConfigurationSheet = configurationDocument.getSheetByName("StyleConfiguration");
var landingConfigurationSheet = configurationDocument.getSheetByName("LandingConfiguration");
var pageConfigurationSheet = configurationDocument.getSheetByName("PageConfiguration");
var informationConfigurationSheet = configurationDocument.getSheetByName("InformationConfiguration");
var menuEntries = [ {name: "Update " + configurationDocument.getName() + " configuration", functionName: "updateConfiguration"},
                    {name: "Display configurationDocumentId", functionName: "displayConfigurationDocumentId"} ];

function onOpen() {
  configurationDocument.addMenu(configurationDocument.getName(), menuEntries);
}

function displayConfigurationDocumentId () {
  Browser.msgBox("The GoogleDocsConfigurationDocumentId for this webapp is " + configurationDocumentId +
                 "The GoogleDocsConfigurationDocumentUrl for this webapp is " + configurationDocumentUrl);
}

// updateConfiguration is a convenience method which sand-boxes the external calls to UrlFetchApp in
// a try-catch block and commits the new configuration to the data-store.
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

// updateSiteConfiguration updates the Site object in the data-store
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
  //headers.t3luxoqcqdwoyagIU4DXvMA = configurationDocumentId;
  headers.daoId = configurationDocumentId;
  var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:siteAttributes};
  if (DEBUG == true) {
    Browser.msgBox("SiteConfiguration Attributes = " + siteAttributes);
  } else {
    if (UrlFetchApp.fetch(siteUrl + siteUpdateUri, advancedArguments).getContentText() != configurationDocumentId) {
      throw "CREATION OF NEW SiteConfiguration OBJECT IN DATASTORE FAILED";
    }
  }
}

// updatesStyleConfiguration updates the Style object in the data-store
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
      headers.daoId = configurationDocumentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:lookAndFeelTypeParameters};
      if (UrlFetchApp.fetch(siteUrl + styleUpdateUri, advancedArguments).getContentText() != configurationDocumentId) {
        throw "CREATION OF NEW StyleConfiguration OBJECT IN APP ENGINE DATASTORE FAILED";
      } 
    }
  }
}

// updateLandingConfiguration creates a new collection of Landing objects in the data-store
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
      headers.daoId = configurationDocumentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:landingParameters};
      if (UrlFetchApp.fetch(siteUrl + landingUpdateUri, advancedArguments).getContentText() != configurationDocumentId) {
        throw "CREATION OF NEW LandingConfiguration OBJECT IN APP ENGINE DATASTORE FOR PAGE FAILED";
      } 
    }
  }
}

// updatePageConfiguration creates a new collection of Page objects in the data-store
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
      headers.daoId = configurationDocumentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:pageAttributes};
      if (UrlFetchApp.fetch(siteUrl + pageUpdateUri, advancedArguments).getContentText() != configurationDocumentId) {
        throw "CREATION OF NEW PageConfiguration OBJECT IN APP ENGINE DATASTORE FOR PAGE " + pageSheets[sheet].getName() + " FAILED";
      }
    }
  }
}

// updateInformationConfiguration creates a new collection of Items in the data-store
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
      headers.daoId = configurationDocumentId;
      var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:informationItemParameters};
      if (UrlFetchApp.fetch(siteUrl + informationUpdateUri, advancedArguments).getContentText() != configurationDocumentId) {
        throw "CREATION OF NEW InformationConfiguration OBJECT IN APP ENGINE DATASTORE FOR PAGE FAILED";
      } 
    }
  }
}

// commitConfigurationChanged persists the newly created object to the data-store and alerts the user
function commitConfigurationChanges() {
  var siteAttributes = "";
  var headers = {};
  headers.gDocsId = configurationDocumentId;
  var advancedArguments = {method:"post", contentType:"application/x-www-form-urlencoded", headers:headers, payload:siteAttributes};
  if (DEBUG != true) {
    if (UrlFetchApp.fetch(siteUrl + commitConfigurationUri, advancedArguments).getContentText() != configurationDocumentId) {
      throw "COMMIT OF NEW CONFIGURATION IN APP ENGINE DATASTORE FOR PAGE FAILED";
    } else {
      Browser.msgBox("Update of " + configurationDocument.getName() + " configuration successfull");
    }
  }
}
  
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