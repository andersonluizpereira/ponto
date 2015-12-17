var aspnm_expandedObjects = new Array();  // array of expanded groups 
var aspnm_expandCount = 0;                // expanded group count 

var aspnm_expandTimerID = 0;          // timer for the current expanding group 
var aspnm_collapseTimerID = 0;        // timer for the current collapsing group 
var aspnm_collapseAllTimerID = 0;     // timer for collapseAll
var aspnm_expandingGroup = '';        // current expanding group name 
var aspnm_collapsingGroup = '';       // current collapsing group name 
var aspnm_restoredGroup = '';         // group that got focus after collapseAll has been called
var aspnm_collapsingAll = false;      // whether the collapseAll command is pending 
var aspnm_curItem = '';               // current menu item 
var aspnm_hideSelectElems = true;     // whether to hide HTML select elements 
var aspnm_shadows = new Array();      // used to hold shadow rectangles
var aspnm_shadowEnabled = false;      // whether shadow is enabled
var aspnm_shadowColor = '#777777';    // shadow color
var aspnm_shadowOffest = 4;           // shadow offset
var aspnm_contextUp = false;          // whether a context menu is up 


// Positions the subgroup for the given item, and initiates its expansion 
function aspnm_itemMsOver(item, subGroup, expandDirection, horAdj, verAdj, expandDelay, effect) 
{
  var newLeft = 0; 
  var newTop = 0; 
  var oItem = document.all[item]; 
  var oSubGroup = document.all[subGroup]; 
    
  if (aspnm_curItem != item)
  {
    aspnm_curItem = item; 
  
    switch (expandDirection)
    {
      case 'belowleft': 
        newLeft = aspnm_pageX(oItem); 
        if (newLeft + oSubGroup.offsetWidth > window.document.body.clientWidth)
          newLeft = aspnm_pageX(oItem) + oItem.offsetWidth - oSubGroup.offsetWidth; 
        newTop = aspnm_pageY(oItem) + oItem.offsetHeight; 
        break; 
      case 'belowright': 
        newLeft = aspnm_pageX(oItem) + oItem.offsetWidth - oSubGroup.offsetWidth; 
        newTop =  aspnm_pageY(oItem) + oItem.offsetHeight; 
        break; 
      case 'aboveleft': 
        newLeft = aspnm_pageX(oItem); 
        newTop =  aspnm_pageY(oItem) - oSubGroup.offsetHeight; 
        break; 
      case 'aboveright': 
        newLeft = aspnm_pageX(oItem) + oItem.offsetWidth - oSubGroup.offsetWidth; 
        newTop =  aspnm_pageY(oItem) - oSubGroup.offsetHeight; 
        break; 
      case 'rightdown': 
        newLeft = aspnm_pageX(oItem) + oItem.offsetWidth; 
        if (newLeft + oSubGroup.offsetWidth > window.document.body.clientWidth)
          newLeft = aspnm_pageX(oItem) - oSubGroup.offsetWidth; 
        newTop = aspnm_pageY(oItem); 
        if (newTop + oSubGroup.offsetHeight > window.document.body.clientHeight)
          newTop = aspnm_pageY(oItem) - oSubGroup.offsetHeight + oItem.offsetHeight; 
        break; 
      case 'rightup': 
        newLeft = aspnm_pageX(oItem) + oItem.offsetWidth; 
        newTop = aspnm_pageY(oItem) - oSubGroup.offsetHeight + oItem.offsetHeight; 
        break; 
      case 'leftdown': 
        newLeft = aspnm_pageX(oItem) - oSubGroup.offsetWidth; 
        newTop = aspnm_pageY(oItem); 
        break; 
      case 'leftup': 
        newLeft = aspnm_pageX(oItem) - oSubGroup.offsetWidth; 
        newTop = aspnm_pageY(oItem) - oSubGroup.offsetHeight + oItem.offsetHeight; 
        break; 
      default: 
        newLeft = aspnm_pageX(oItem) + oItem.offsetWidth; 
        newTop = aspnm_pageY(oItem); 
        break; 
    }  
    newLeft += horAdj; 
    newTop += verAdj; 
    if (newTop < 0) newTop = 0; 
    if (newLeft < 0) newLeft = 0; 
    
    oSubGroup.style.left = newLeft + 'px'; 
    oSubGroup.style.top = newTop + 'px'; 

    aspnm_startExpand(subGroup, expandDelay, effect); 
  }
}

// If the mouse pointer is not over the given item or its subGroup, 
// this function initiates the collapse of the subGroup. 
function aspnm_itemMsOut(item, group, subGroup, expandDelay, effect)
{
  if ((!(aspnm_isMouseOnObject(item))) && subGroup) 
    if (!(aspnm_isMouseOnObject(subGroup)))
    {
      aspnm_curItem = ''; 
      aspnm_startCollapse(subGroup, expandDelay, effect);
    }  
}

// This event handler is only called on expandable groups. If collapseAll is pending, it sets the 
// global variable aspnm_restoredGroup to the given group, so that the group and its parent groups 
// are not collapsed. It also stops the collapse if it has been issued for this group. 
function aspnm_groupMsOver(group)
{
  if (aspnm_collapsingAll) aspnm_restoredGroup = group; 

  if (aspnm_collapsingGroup == group) 
  {
    aspnm_stopCollapse(); 
    aspnm_stopExpand(); 
  }
}

// If the mouse pointer is on the given group, its subGroup, or its parent item this function 
// does nothing. If the pointer is over the parent group, but outside the parent item, then it
// initiates the collapse of itself and its subGroup (if any). 
// Otherwise, the pointer is outside the menu structure and it initiates the collapse of all 
// expanded objects. 
function aspnm_groupMsOut(group, parentItem, parentGroup, expandDelay, effect)
{ 
  if (!(aspnm_isMouseOnObject(group)))
  {
    aspnm_curItem = ''; 

    var subGroup = aspnm_expandedObjects[aspnm_expandCount]; 
    if (subGroup == group) subGroup = null; 

    if (parentItem == null && parentGroup == null && !(aspnm_isMouseOnObject(group)))
      aspnm_startCollapseAll(expandDelay, effect);     
    else if (aspnm_isMouseOnObject(group) || aspnm_isMouseOnObject(subGroup) || aspnm_isMouseOnObject(parentItem))
      ; // do nothing 
    else if (aspnm_isMouseOnObject(parentGroup))
    {
      aspnm_startCollapse(group, expandDelay, effect); 
      aspnm_startCollapse(subGroup, expandDelay, effect); 
    }
    else
      aspnm_startCollapseAll(expandDelay, effect);     
  }
}

// Expand/collapse timer functinos ----------------------------------------------------------------

// Initiates the expand of the given group. 
function aspnm_startExpand(group, interval, effect)
{
  if (group == aspnm_collapsingGroup) aspnm_stopCollapse(); 
  if (group != aspnm_expandingGroup) aspnm_stopExpand();  
  
  aspnm_restoredGroup = group; 
  
  aspnm_expandingGroup = group; 
  if (group) group += '.id'; 
  if (effect) effect = "'" + effect + "'";  
  aspnm_expandTimerID = setTimeout('aspnm_expand(' + group + ', ' + effect + ')', interval); 
}

// Initiates the collapse of the given group. 
function aspnm_startCollapse(group, interval, effect)
{
  if (group == aspnm_expandingGroup) aspnm_stopExpand(); 

  if (group) 
    if (document.all[group].style.visibility == 'visible') 
    {
      aspnm_collapsingGroup = group; 
      group += '.id'; 
      if (effect) effect = "'" + effect + "'";  
      aspnm_collapseTimerID = setTimeout('aspnm_collapse(' + group + ', ' + effect + ')', interval); 
    }  
}

// Initiates the collapse of all expanded objects. 
function aspnm_startCollapseAll(interval, effect)
{
  aspnm_stopCollapse(); 
  aspnm_stopExpand(); 
  aspnm_stopCollapseAll(); 

  aspnm_collapsingAll = true; 
  if (effect) 
  {
    effect = "'" + effect + "'";  
    aspnm_collapseAllTimerID = setTimeout('aspnm_collapseAll(' + effect + ')', interval); 
  }
  else
    aspnm_collapseAllTimerID = setTimeout('aspnm_collapseAll(null)', interval); 
}

// Stops the expand of the currently expanding group. 
function aspnm_stopExpand()
{
  clearTimeout(aspnm_expandTimerID); 
  aspnm_expandingGroup = ''; 
}

// Stops the collapse of the currently collapsing group. 
function aspnm_stopCollapse()
{
  clearTimeout(aspnm_collapseTimerID); 
  aspnm_collapsingGroup = ''; 
}

// Stops the collapse of all currently expanding objects. 
function aspnm_stopCollapseAll()
{
  clearTimeout(aspnm_collapseAllTimerID); 
  aspnm_restoredGroup = '';
}


// Core functions ---------------------------------------------------------------------------------

// Expands the given menu group 
function aspnm_expand(group, effect)
{
  if (document.all[group].style.visibility != 'visible')
  {
    aspnm_hideSelectElements(group); 
    if (effect) 
    {
      document.all[group].style.filter = effect; 
      document.all[group].filters[0].Apply(); 
    }  
    document.all[group].style.visibility = 'visible'; 
    aspnm_makeDropShadow(group); 
    if (effect) document.all[group].filters[0].Play(); 
    aspnm_expandCount++; 
    aspnm_expandedObjects[aspnm_expandCount] = group; 
  }  
}


// Collapses the given menu group 
function aspnm_collapse(group, effect)
{
  if (group) 
  {
    if (document.all[group].style.visibility != 'hidden')
    {
      if (effect)
      {
        document.all[group].style.filter = effect; 
        document.all[group].filters[0].Apply(); 
      }
      document.all[group].style.visibility = 'hidden';     
      if (effect) document.all[group].filters[0].Play(); 
      aspnm_expandCount--; 
      aspnm_clearDropShadow(group); 
    }      
  }
  if (!(aspnm_contextUp) && aspnm_expandCount == 0) 
    aspnm_restoreSelectElements(); 
}

// Collapses all expanded menu groups 
function aspnm_collapseAll(effect)
{
  for (var i = aspnm_expandCount; i >= 1; i--)
  {
    if (aspnm_expandedObjects[i] == aspnm_restoredGroup) break; 

    if (effect)
    {
      document.all[aspnm_expandedObjects[i]].style.filter = effect; 
      document.all[aspnm_expandedObjects[i]].filters[0].Apply(); 
    }
    document.all[aspnm_expandedObjects[i]].style.visibility = 'hidden';
    aspnm_clearDropShadow(aspnm_expandedObjects[i]); 
    if (effect) document.all[aspnm_expandedObjects[i]].filters[0].Play(); 
  }

  aspnm_collapsingAll = false; 
  aspnm_expandCount = i;
  aspnm_restoredGroup = ''; 
  if (!(aspnm_contextUp) && aspnm_expandCount == 0) 
    aspnm_restoreSelectElements(); 
}

// Hides all menu groups prior to calling ClientSideOnClick event handler
function aspnm_hideAllGroups()
{
  aspnm_curItem = ''; 
  aspnm_restoredGroup = ''; 
  aspnm_collapseAll(null); 
}

// Utilities --------------------------------------------------------------------------------------

// Updates menu item class, left icon, and right icon 
function aspnm_updateCell(Element, NewClassName, LeftImage, LeftImageSrc, RightImage, RightImageSrc, direction)
{  
  if (direction == 'out' && aspnm_isMouseOnObject(Element))    
    ;   
  else  
  {    
    if (Element != null & NewClassName != '') document.all[Element].className = NewClassName;
    if (LeftImage != null  && LeftImageSrc != '') document.images[LeftImage].src = LeftImageSrc;     
    if (RightImage != null && RightImageSrc != '') document.images[RightImage].src = RightImageSrc;   
  }
}

// Determines whether the mouse pointer is currently over the given object 
function aspnm_isMouseOnObject(objName)
{
  if (objName)
  {
    var objLeft = aspnm_pageX(document.all[objName]) - window.document.body.scrollLeft + 1; 
    var objTop = aspnm_pageY(document.all[objName]) - window.document.body.scrollTop + 1; 
    var objRight = objLeft + document.all[objName].offsetWidth - 1; 
    var objBottom = objTop + document.all[objName].offsetHeight - 1;
  
    if ((event.x > objLeft) && (event.x < objRight) && 
        (event.y > objTop) && (event.y < objBottom))
      return true; 
    else  
      return false; 
  }
  else
    return false; 
}

// Calculates the absolute page x coordinate of any element
function aspnm_pageX(element)
{
  var x = 0;
  do 
  {
    
    if (element.style.position == 'absolute') 
    {
      return x + element.offsetLeft; 
    }
    else
    {
      x += element.offsetLeft;
      if (element.offsetParent) 
        if (element.offsetParent.tagName == 'TABLE') 
          if (parseInt(element.offsetParent.border) > 0)
          {
            x += 1; 
          }
    }
  }
  while ((element = element.offsetParent));
  return x; 
}

// Calculates the absolute page y coordinate of any element
function aspnm_pageY(element)
{
  var y = 0;
  do 
  {
    if (element.style.position == 'absolute') 
    {
      return y + element.offsetTop; 
    }
    else
    {
      y += element.offsetTop;
      if (element.offsetParent) 
        if (element.offsetParent.tagName == 'TABLE') 
          if (parseInt(element.offsetParent.border) > 0)
          {
            y += 1; 
          }
    }
  }
  while ((element = element.offsetParent));
  return y; 
}


// Hides HTML select elements that are overlapping the given menu group 
function aspnm_hideSelectElements(group)
{
  if (document.getElementsByTagName) 
  {
    var arrElementsAll = document.getElementsByTagName('select'); 
	var arrElements    = document.getElementsByTagName('select');
		
    if (aspnm_hideSelectElems) 
    {
		for (var i = 0; i < arrElementsAll.length; i++) 
			if (aspnm_objectsOverlapping(document.all[group], arrElementsAll[i]))
				arrElementsAll[i].style.visibility = 'hidden';          
		
		for (var i = 0; i < arrElements.length; i++) {
			arrElements[i].style.visibility    = 'hidden';
		}
		
    }
  }
}

// Restores all HTML select elements on the page 
function aspnm_restoreSelectElements()
{
	if (document.getElementsByTagName) 
	{
		var arrElementsAll = document.getElementsByTagName('select'); 
		var arrElements    = document.getElementsByTagName('select');
		
		if (aspnm_hideSelectElems) {
			for (var i = 0; i < arrElementsAll.length; i++) {
				arrElementsAll[i].style.visibility = 'visible';
			} 
			for (var i = 0; i < arrElements.length; i++) {
				arrElements[i].style.visibility    = 'visible';
			}
		}
	}
}

// Whether the given objects are overlapping 
function aspnm_objectsOverlapping(obj1, obj2)
{
  var result = true; 
  var obj1Left = aspnm_pageX(obj1) - window.document.body.scrollLeft; 
  var obj1Top = aspnm_pageY(obj1) - window.document.body.scrollTop; 
  var obj1Right = obj1Left + obj1.offsetWidth; 
  var obj1Bottom = obj1Top + obj1.offsetHeight;
  var obj2Left = aspnm_pageX(obj2) - window.document.body.scrollLeft; 
  var obj2Top = aspnm_pageY(obj2) - window.document.body.scrollTop; 
  var obj2Right = obj2Left + obj2.offsetWidth; 
  var obj2Bottom = obj2Top + obj2.offsetHeight;
  
  if (obj1Right <= obj2Left || obj1Bottom <= obj2Top || 
      obj1Left >= obj2Right || obj1Top >= obj2Bottom) 
    result = false; 
  return result; 
}

// Creates a drop shadow for an object 
function aspnm_makeDropShadow(objName)
{
  if (aspnm_shadowEnabled) 
  {
    aspnm_shadows[objName] = new Array(); 
	  for (var i = aspnm_shadowOffest; i > 0; i--)
	  {
	    var obj = document.all[objName]; 
		  var rect = document.createElement('div');
		  var rs = rect.style
		  rs.position = 'absolute';
		  rs.left = (obj.style.posLeft + i) + 'px';
		  rs.top = (obj.style.posTop + i) + 'px';
		  rs.width = obj.offsetWidth + 'px';
		  rs.height = obj.offsetHeight + 'px';
		  rs.zIndex = obj.style.zIndex - i;
		  rs.backgroundColor = aspnm_shadowColor;
		  var opacity = 1 - i / (i + 1);
		  rs.filter = 'alpha(opacity=' + (100 * opacity) + ')';
		  obj.insertAdjacentElement('afterEnd', rect);
		  aspnm_shadows[objName][aspnm_shadows[objName].length] = rect; 
	  }
	}
}

// Clears the drop shadow for the given object 
function aspnm_clearDropShadow(objName)
{
  if (aspnm_shadowEnabled) 
  {
    var curShadow; 
    for (var i = 0; i < aspnm_shadows[objName].length; i++)
    {
      curShadow = aspnm_shadows[objName][i]; 
      curShadow.style.filter = 'alpha(opacity=0)'; 
      curShadow.removeNode(true); 
    }
  }  
}

// Positions the menu based on the alignment, offsetX, and offsetY properties
function aspnm_positionMenu(menu, alignment, offsetX, offsetY)
{
  var scrlLeft = 0; 
  var scrlTop = 0;
  var clientW = window.document.body.clientWidth; 
  var clientH = window.document.body.clientHeight; 
  var menuWidth = menu.offsetWidth; 
  var menuHeight = menu.offsetHeight; 
  var newLeft = 0; 
  var newTop = 0; 

  switch (alignment)
  {
    case 'topleft': 
      newLeft = scrlLeft;
      newTop = scrlTop;
      break; 
    case 'topmiddle': 
      newLeft = (clientW - menuWidth) / 2 + scrlLeft;
      newTop = scrlTop;
      break; 
    case 'topright': 
      newLeft = clientW + scrlLeft - menuWidth;
      newTop = scrlTop;
      break; 
    case 'bottomleft': 
      newLeft = scrlLeft;
      newTop = clientH + scrlTop - menuHeight;
      break; 
    case 'bottommiddle': 
      newLeft = (clientW - menuWidth) / 2 + scrlLeft;
      newTop = clientH + scrlTop - menuHeight;
      break; 
    case 'bottomright': 
      newLeft = clientW + scrlLeft - menuWidth;
      newTop = clientH + scrlTop - menuHeight;
      break; 
    default: 
      newLeft = clientW + scrlLeft;
      newTop = clientH + scrlTop;
	  break; 
  }    
  
  newLeft += offsetX; 
  newTop += offsetY; 
  menu.style.left = newLeft; 
  menu.style.top = newTop; 
}
