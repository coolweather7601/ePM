/* Define the site identitier */
_page.siteIdentifier = 'APK ePM <font color="red">(Production site)</font>'
_page.hideLeftNavigation = true;

/* Define the site menu structure using an associative array */

//_page.items["0"]        = new _Item("List","../others/listIndex.aspx")

_page.items["0"]        = new _Item("List","../eForm/List.aspx")
_page.items["1"]        = new _Item("Sheet_Category","../eForm/Sheet_Category.aspx")
_page.items["2"]        = new _Item("ACS_Manage","../others/ACS_Manage.aspx")

/* Define the site menu structure using an associative array */
_page.sites[0] = new _Site("Semiconductors", "http://twgkhhpsk1ms011 ")
_page.sites[1] = new _Site("APK","http://twgkhhpsk1ms011/imo/backend/kaohsiung/")


/* This table of content is required for the Autonomy crawler 
<meta name="robots" content="noindex, follow" />
<a href="index.html"></a>
*/
