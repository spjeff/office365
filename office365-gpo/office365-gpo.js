/* Office365 - Group Policy

 - hide Site Setting links
 - hide Site Features
 - hide Web Features
 - hide Custom Permission Levels
 
   last updated 03-12-18
*/

(function() {
    //wait for DOM element to appear
    function waitDom(id, fn) {
        var checkExist = setInterval(function() {
            if (document.getElementById(id)) {
                clearInterval(checkExist);
                fn();
            }
        }, 250);
    }

    //Inject CSS
    function injectCss(css) {
        var style = document.createElement("style");
        style.type = "text/css";
        style.innerHTML = style;
        document.body.appendChild(style);
    }

    //hide site feature
    function hideSPFeature(name) {
        var headers = document.querySelectorAll('h3.ms-standardheader');
        if (headers) {
            for (var i = 0; i < headers.length; i++) {
                if (headers[i].innerText.indexOf(name) >= 0) {
                    headers[i].parentNode.parentNode.parentNode.parentNode.parentNode.parentNode.style.display = 'none';
                }
            }
        }
    }

    //remove alternating row color
    function hideAltRowColor() {
        var className = 'ms-featurealtrow';
        var rows = document.querySelectorAll('td.' + className);
        if (rows) {
            for (var i = 0; i < rows.length; i++) {
                var el = rows[i];
                if (el.classList) {
                    el.classList.remove(className);
                } else {
                    el.className = el.className.replace(new RegExp('(^|\\b)' + className.split(' ').join('|') + '(\\b|$)', 'gi'), ' ');
                }

            }
        }
    }

    //menu - hide Web Features
    function menuWebFeatures() {
        //hide rows
        var features = ['Access App',
            'Announcement Tiles',
            'Community Site Feature',
            'Duet Enterprise - SAP Workflow',
            'Duet Enterprise Reporting',
            'Duet Enterprise Site Branding',
            'External System Events',
            'Getting Started with Project Web App',
            'Hold',
            'Minimal Download Strategy',
            'Offline Synchronization for External Lists',
            'Project Functionality',
            'Project Proposal Workflow',
            'Project Web App Connectivity',
            'SAP Workflow Web Parts',
            'Search Config Data Content Types',
            'Search Config Data Site Columns',
            'Search Config List Instance Feature',
            'Search Config Template Feature',
            'Site Feed',
            'SharePoint Server Publishing'
        ];
        features.forEach(function(feature, i) {
            hideSPFeature(feature);
        });

        //hide row background color
        hideAltRowColor();
    }

    //menu - hide Site Features
    function menuSiteFeatures() {
        //hide rows
        var features = ['Content Type Syndication Hub',
            'Custom Site Collection Help',
            'Cross-Site Collection Publishing',
            'Duet End User Help Collection',
            'Duet Enterprise Reports Content Types',
            'In Place Records Management',
            'Library and Folder Based Retention',
            'Limited-access user permission lockdown mode',
            'Project Server Approval Content Type',
            'Project Web App Permission for Excel Web App Refresh',
            'Project Web App Ribbon',
            'Project Web App Settings',
            'Publishing Approval Workflow',
            'Reports and Data Search Support',
            'Sample Proposal',
            'Search Engine Sitemap',
            'SharePoint 2007 Workflows',
            'SharePoint Server Publishing Infrastructure',
            'Site Policy',
            'Workflows'
        ];
        features.forEach(function(feature, i) {
            hideSPFeature(feature);
        });

        //hide row background color
        hideAltRowColor();
    }

    //menu - hide Settings Links
    function menuSettings() {
        //hide links
        var links = ['ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_AuditSettings',
            'ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SharePointDesignerSettings',
            'ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_PolicyPolicies',
            'ctl00_PlaceHolderMain_SiteAdministration_RptControls_PolicyPolicyAndLifecycle',
            'ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_HubUrlLinks',
            'ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_Portal',
            'ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_HtmlFieldSecurity',
            'ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SearchConfigurationImportSPSite',
            'ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SearchConfigurationExportSPSite'
        ];
        links.forEach(function(id, i) {
            var el = document.getElementById(id);
            if (el) {
                el.style.display = 'none';
            }
        });

        // Change Owner link
        // find group
        var match;
        var section = document.querySelectorAll('h3.ms-linksection-title');
        if (section) {
            for (var i = 0; i < section.length; i++) {
                var el = section[i];
                if (el.innerHTML.indexOf("Users and Permissions") > 0) {
                    match = el.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode;
                }
            }
        }
        if (match) {
            //append new child link
            var group = match.querySelector("ul");
            var li = document.createElement("li");
            li.innerHTML = '<a title="Change site owner." href="setrqacc.aspx?type=web">Change Site Owner</a>';
            group.appendChild(li);
        }
    }
	
	//toolbar - permission levels
	function userSettings() {
		var el = document.getElementById('Ribbon.Permission.Manage-LargeMedium-0-0');
		if (el) {
			el.style.display = 'none';
		}
	}
	
	//body - permission levels
	function roleSettings() {
		var el = document.getElementById('DeltaPlaceHolderMain');
		if (el) {
			el.innerHTML = '<h2>Disabled by SharePoint Support team.</h2>';
		}
    }
    
    //access request - checkboxes
    function accessSettings() {
        var css = '#ctl00_PlaceHolderMain_ctl00_chkMembersCanShare,label[for="ctl00_PlaceHolderMain_ctl00_chkMembersCanShare"],#ctl00_PlaceHolderMain_ctl00_chkMembersCanAddToGroup,label[for="ctl00_PlaceHolderMain_ctl00_chkMembersCanAddToGroup"] {display:none;}';
        injectCss(css);
    }

    //wait until document ready http://youmightnotneedjquery.com/
    function ready(fn) {
        if (document.readyState != 'loading') {
            fn();
        } else {
            document.addEventListener('DOMContentLoaded', fn);
        }
    }

    //URL contains expression
    function urlContains(expr) {
        return document.location.href.toLowerCase().indexOf(expr.toLowerCase()) > 0;
    }

    //core logic
    function main() {
        if (!urlContains('skip')) {
            //Web Features
            if (urlContains('ManageFeatures.aspx') && !urlContains('Scope=Site')) {
                waitDom('DeltaPlaceHolderMain', menuWebFeatures);
            }

            //Site Features
            if (urlContains('ManageFeatures.aspx?Scope=Site')) {
                waitDom('DeltaPlaceHolderMain', menuSiteFeatures);
            }

            //Site Settings
            if (urlContains('settings.aspx')) {
                waitDom('DeltaPlaceHolderMain', menuSettings);
            }
			
			//User
            if (urlContains('user.aspx')) {
                waitDom('Ribbon.Permission.Manage', userSettings);
            }
			
			//Role
            if (urlContains('role.aspx')) {
                waitDom('DeltaPlaceHolderMain', roleSettings);
            }

            //Access Request
            if (urlContains('setrqacc.aspx')) {
                waitDom('DeltaPlaceHolderMain', accessSettings);
            }
            
        }
    }


    //initialize
    ready(main);
})();