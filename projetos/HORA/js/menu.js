// creates a new drop-down menu
DropMenu = function(id)
{
	var me = this;
	
	this.element = document.all ? document.all[id] : document.getElementById(id);
	this.fixHover = true; //document.all && navigator.appVersion.indexOf("MSIE 7.0") < 0;
	this.items = new Array();
	
	if (document.getElementsByTagName)
	{
		me.initialize();
	}
}

// initializes the menu
DropMenu.prototype.initialize = function()
{
	var root = this.element;
	var list = root.getElementsByTagName("li");
	
	// discover menu items
	for (var i = 0; i < list.length; i++)
	{
		var item = list[i];
		
		if (item.parentNode != null && item.parentNode.parentNode == root)
		{
			this.items.push(new DropMenuItem(item, this));
		}
	}
	
	// fix menu z-order when popping over form elements in IE
	if (document.all)
	{
		var subs = root.getElementsByTagName("ul");
		
		for (var i = 0; i < subs.length; i++)
		{
			var frameMenu = subs[i];
			var frame = frameMenu.appendChild(document.createElement("iframe"));
								
			frame.frameBorder = "0";
			frame.scrolling = "no";
			frame.src = "about:blank";
			frame.style.filter = "progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)";
			frame.style.left = "0px";
			frame.style.height = frameMenu.offsetHeight + "px";
			frame.style.position = "absolute";
			frame.style.top = "0px";
			frame.style.visibility = "hidden";
			frame.style.width = frameMenu.offsetWidth + "px";
			frame.style.zIndex = "-1";
			frameMenu.style.zIndex = "101";
		}
	}

	return this;
}

// creates a new item
DropMenuItem = function(element, menu, parent)
{
	var me = this;
	
	this.childItems = new Array();
	this.element = element;
	this.menu = menu;
	this.parent = parent;
	
	// hover fix for IE6
	if (this.menu.fixHover)
	{
		var instance = this;
		
		element.onmouseover = function(){instance.show();}
		element.onmouseout = function(){instance.timeout = window.setTimeout(function(){instance.hide();}, 500);}
		element.toggleSubmenus = function()
		{
			var frames = this.getElementsByTagName("iframe");
			
			for (var i = 0; i < frames.length; i++)
			{
				frames[i].style.visibility = frames[i].parentNode.currentStyle.visibility;

				if (!frames[i].positioned)
				{
					frames[i].style.height = frames[i].parentNode.offsetHeight + "px";
					frames[i].style.width = frames[i].parentNode.offsetWidth + "px";
					frames[i].positioned = true;
				}
			}
		}
	}
	
	// initialize the menu
	if (document.getElementsByTagName)
	{
		me.initialize();
	}
}

// gets the nesting depth of the menu
DropMenuItem.prototype.getDepth = function()
{
	if (this.parent != null)
	{
		return this.parent.getDepth() + 1;
	}
	
	return 0;
}

// gets the width of the item
DropMenuItem.prototype.getWidth = function()
{
	if (this.element != null)
	{
		return Math.max(this.element.offsetWidth, this.element.parentNode.offsetWidth);
	}
	
	return 0;
}

// initializes the item
DropMenuItem.prototype.initialize = function()
{
	var root = this.element;
	var list = root.getElementsByTagName("li");
	var size = this.getWidth();
	
	// add child items to the menu
	for (var i = 0; i < list.length; i++)
	{
		var listItem = list[i];
		var menuItem = null;
		
		if (listItem.parentNode != null && listItem.parentNode.parentNode == root)
		{
			menuItem = new DropMenuItem(listItem, this.menu, this);
			size = Math.max(menuItem.element.offsetWidth, size);
			
			this.childItems.push(menuItem);
		}
	}
	
	if (this.parent != null)
	{
		this.element.parentNode.style.width = size + "px";
		this.element.style.width = "100%";
		
		if (this.getDepth() == 1)
		{
			var parent = this.parent;
			var parentOffset = size;
		
			while (parent != null)
			{
				parentOffset += parent.element.offsetLeft;
				parent = parent.parent;
			}
		
			this.element.parentNode.style.left = Math.min(this.menu.element.offsetWidth - parentOffset - 2, 0) + "px";
		}
	}
}

DropMenuItem.prototype.hide = function()
{
	if (this.element)
	{
		this.element.className = this.element.className.replace(/\s*hover/gi, "");
		this.element.toggleSubmenus();
	}
}

DropMenuItem.prototype.show = function()
{
	var items = this.parent != null ? this.parent.childItems : this.menu.items;
	
	// set the classname and toggle iframe fix
	if (this.element)
	{
		this.element.className = this.element.className.replace(/\s*hover/gi, "") + " hover";
		this.element.toggleSubmenus();
	}
	
	// clear the hide timer	
	if (this.timeout)
	{
		window.clearTimeout(this.timeout);
	}

	// hide any other open menus
	for (var i = 0; i < items.length; i++)
	{
		if (items[i] != this)
		{
			items[i].hide();
		}
	}
}