function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Refresh")
  .addItem("Refresh page", "refreshPage")
  .addItem("Refresh all pages", "refreshAllPages")
  .addToUi();
}

function refreshPage() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertColumnBefore(1);
  sheet.deleteColumn(1);
}

function refreshAllPages() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  for(var i = 0; i < sheets.length; i++) {
    sheets[i].insertColumnBefore(1);
    sheets[i].deleteColumn(1);
  }
}

/******************The following functions were created for the buttons on the sheet.******************/

// Since there can only be one filter in a sheet at a time, resetFilters resets any pre-existing filter.
function resetFilters() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getFilter().remove();
}
// All members of the discord server are listed starting from A13 in the sheet. This function 
// checks that there are no active filters in the sheet and then returns the data range A13:J83.
function rangeHelper() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getFilter() != null) sheet.getFilter().remove();
  return sheet.getRange(13, 1, 70, 10);
}
// Filters the data to show only mains.
function showMainsOnly() {
  var range = rangeHelper()
  var filterCriteria = SpreadsheetApp.newFilterCriteria().whenTextContains('Main').build();
  return range.createFilter().setColumnFilterCriteria(4, filterCriteria).getRange();
}
// Filters the data to show only alts.
function showAltsOnly() {
  var range = rangeHelper()
  var filterCriteria = SpreadsheetApp.newFilterCriteria().whenTextContains('Alt').build();
  return range.createFilter().setColumnFilterCriteria(4, filterCriteria).getRange();
}
// Filters the data to show only clan members.
function showMembersOnly() {
  var range = rangeHelper()
  var filterCriteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('Member').build();
  return range.createFilter().setColumnFilterCriteria(3, filterCriteria).getRange();
}
// Filters the data to show only new clan members.
function showNewbiesOnly() {
  var range = rangeHelper()
  var filterCriteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('Newbie').build();
  return range.createFilter().setColumnFilterCriteria(3, filterCriteria).getRange();
}
// Filters the data to show only ex clan members.
function showExMembersOnly() {
  var range = rangeHelper()
  var filterCriteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo('Ex-member').build();
  return range.createFilter().setColumnFilterCriteria(3, filterCriteria).getRange();
}
// Sorts the dataset by order: Newbies > Members > Ex-members.
function sortByMembership() {
  var range = rangeHelper()
  return range.sort({column: 3, ascending: false});
}
// Sorts the dataset by order: Main > Alt.
function sortByMain() {
  var range = rangeHelper()
  return range.sort({column: 4, ascending: false});
}
// Sorts the dataset by number of characters they have in the clan in descending order.
function sortByCharNumber() {
  var range = rangeHelper()
  return range.sort({column: 6, ascending: false});
}
// Sorts the dataset by character name in ascending order.
function sortByCharName() {
  var range = rangeHelper()
  return range.sort({column: 1, ascending: true});
}
// Sorts the dataset by how long they have been in the clan (or rather the discord server).
function sortByDays() {
  var range = rangeHelper()
  return range.sort({column: 9, ascending: true});
}

/******************The following functions were created for viewing BnS character stats.******************/

// Returns an object to a desired item in the form: {name: accesoryName, image: imageSrc}.
// See bnsLoadImage and bnsLoadName for reference.
function __xmlParser(doc, item) {
  var result = {};
  var empty = false;
  
  // Default CSS class is weapon.
  if(item == "weapon") {
    var cursor = doc.slice(doc.indexOf('<div class="wrapWeapon">')+'<div class="wrapWeapon">'.length);
    
    if(cursor.indexOf('<img src="') > cursor.indexOf('<div class="wrapAccessory'))
      empty = true;
    
    cursor = cursor.slice(cursor.indexOf("<img src=") + '<img src="'.length);
    
    result.image = cursor.substr(0, cursor.indexOf('"'));
    
    cursor = cursor.slice(cursor.indexOf('<div class="name">') + '<div class="name">'.length);
    cursor = cursor.slice(cursor.indexOf('>') + 1);
    
    result.name = cursor.substr(0, cursor.indexOf('<'));
  }
  else {
    // Translating to the names of the CSS classes.
    var hashing = {
      "Ring":"ring",
      "Earring":"earring",
      "Necklace":"necklace",
      "Bracelet":"bracelet",
      "Belt":"belt",
      "Gloves":"gloves",
      "Soul":"soul",
      "Heart":"soul-2",
      "Pet":"guard",
      "Talisman":"nova",
      "Soulbadge":"singongpae",
      "Mysticbadge":"rune"
    };
    
    var cursor = doc.slice(doc.indexOf('<div class="wrapAccessory '+hashing[item]+'">')+('<div class="wrapAccessory '+hashing[item]+'">').length);
     
    // Checking to see if the selected image actually belongs to the desired item.
    // In case there is nothing equipped in its slot, the script would select
    // the image of the next item. This condition checks for that.
    if(cursor.indexOf('<img src="') > cursor.indexOf('<div class="wrapAccessory'))
      empty = true;
    
    cursor = cursor.slice(cursor.indexOf('<img src="') + '<img src="'.length);
    
    result.image = cursor.substr(0, cursor.indexOf('"'));  
    
    cursor = cursor.slice(cursor.indexOf('<div class="name">') + '<div class="name">'.length);
    cursor = cursor.slice(cursor.indexOf('>')+1);
    
    result.name = cursor.substr(0, cursor.indexOf('<'));
    
  }
  
  if(empty) {
    result.image = "https://i.imgur.com/DrJhfd4.png";
    result.name = "EMPTY";
  }
  
  return result;
  
}
// Fetches the gear icon for the item (if equipped by the character) specified in the parameter.
// @Params:
// charname = name of the character ingame.
// item = Weapon, Ring, Earring, Necklace, Bracelet, Belt, Gloves, Soul, Heart, Pet, Talisman, Soulbadge, Mysticbadge.
function loadGearImage(charname, item) {
  return __xmlParser(
    UrlFetchApp.fetch("http://eu-bns.ncsoft.com/ingame/bs/character/data/equipments?c="+charname).getContentText(),
    item
  ).image;
}
// Fetches the gear name for the item (if equipped by the character) specified in the parameter.
// @Params:
// charname = name of the character ingame.
// item = Weapon, Ring, Earring, Necklace, Bracelet, Belt, Gloves, Soul, Heart, Pet, Talisman, Soulbadge, Mysticbadge.
function loadGearName(charname, item) {
  return __xmlParser(
    UrlFetchApp.fetch("http://eu-bns.ncsoft.com/ingame/bs/character/data/equipments?c="+charname).getContentText(),
    item
  ).name;
}
// Fetches the class icon of the character specified in the parameter.
// @Params:
// charname = name of the character ingame.
function loadClassIcon(charname) {
  var doc = UrlFetchApp.fetch("http://eu-bns.ncsoft.com/ingame/bs/character/profile?c="+charname).getContentText();
  doc = doc.slice(doc.indexOf('<div class="classThumb"><img src="') + '<div class="classThumb"><img src="'.length);
  
  return doc.substr(0, doc.indexOf('"'));
}
// Fetches the account name of the character specified in the parameter.
// @Params:
// charname = name of the character ingame.
function loadAccountName(charname) {
   var doc = UrlFetchApp.fetch("http://eu-bns.ncsoft.com/ingame/bs/character/profile?c="+charname).getContentText();
  doc = doc.slice(doc.indexOf('<dt><a href="#">') + '<dt><a href="#">'.length);
  
  return doc.substr(0, doc.indexOf('<'));
}
// Fetches the stats of the character specified in the parameter.
// @Params: 
// charname = name of the character ingame. 
// stat = Attack Power, Crit Rate, Crit Damage, Accuracy, Mystic, HP
function loadStats(charname, stat) {
  var response = UrlFetchApp.fetch("http://eu-bns.ncsoft.com/ingame/bs/character/data/abilities.json?c=" + charname);
  
  var charStats = JSON.parse(response.getContentText()).records.total_ability;
  
  switch(stat) {
    case "Attack Power":
      return charStats.attack_power_value;
      break;
    case "Crit Rate":
      return charStats.attack_critical_value + " ("+charStats.attack_critical_rate+"%)";
      break;
    case "Crit Damage":
      return charStats.attack_critical_damage_value + " ("+charStats.attack_critical_damage_rate+"%)";
      break;
    case "Accuracy":
      return charStats.attack_hit_value + " ("+charStats.attack_hit_rate+"%)";
      break;
    case "Mystic":
      return charStats.attack_attribute_value + " ("+charStats.attack_attribute_rate+"%)";
      break;
    case "HP":
      return charStats.int_max_hp;
      break;
    default:
      return "BAD PARAMETER";
  }
}
