/*
Office365 - Group Policy

 * hide Site Setting links
 * hide Site Features
 * hide Web Features
 
 * last updated 08-01-16
*/

(function() {
    //hide site feature
    function hideSPFeature(name) {
        var el = document.querySelector('h3.ms-standardheader:contains("' + name + '")').parentNode.parentNode.parentNode.parentNode.parentNode.parentNode;
        el.parentNode.removeChild(el);
    };

    //URL contains expression
    function urlContains(expr) {
        return document.location.href.toLowerCase().indexOf(expr.toLowerCase()) > 0;
    };

    //remove alternating row color
    function hideAltRowColor() {
        var rows = document.querySelectorAll('td.ms-featurealtrow')
        rows.forEach(function(el, i) {
            var className = 'ms-featurealtrow';
            if (el.classList) {
                el.classList.remove(className);
            } else {
                el.className = el.className.replace(new RegExp('(^|\\b)' + className.split(' ').join('|') + '(\\b|$)', 'gi'), ' ');
            }
        });
    };

    //core logic
    function main() {
        if (!urlContains('skip')) {
            //Web Features
            if (urlContains('ManageFeatures.aspx') && !urlContains('Scope=Site')) {
                //hide rows
                var features = ['Access App',
                    'Announcement Tiles',
                    'Community Site Feature',
                    'Duet Enterprise - SAP Workflow',
                    'Duet Enterprise Reporting',
                    'Duet Enterprise Site Branding',
                    'Getting Started with Project Web App',
                    'Minimal Download Strategy',
                    'Project Functionality',
                    'Project Proposal Workflow',
                    'Project Web App Connectivity',
                    'SAP Workflow Web Parts',
                    'SharePoint Server Publishing'
                ];
                features.forEach(function(feature, i) {
                    hideSPFeature(feature);
                });

                //hide row background color
                hideAltRowColor();
            }

            //Site Features
            if (urlContains('ManageFeatures.aspx?Scope=Site')) {
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

            //Site Settings
            if (urlContains('settings.aspx')) {
                //hide links
                var links = ['#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_AuditSettings',
                    '#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SharePointDesignerSettings',
                    '#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_PolicyPolicies',
                    '#ctl00_PlaceHolderMain_SiteAdministration_RptControls_PolicyPolicyAndLifecycle',
                    '#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_HubUrlLinks',
                    '#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_Portal',
                    '#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_HtmlFieldSecurity',
                    '#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SearchConfigurationImportSPSite',
                    '#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SearchConfigurationExportSPSite'
                ];
                links.forEach(function(el, i) {
                    el.parentNode.removeChild(el);
                });

                // Change Owner link
                // find group
                var match;
                document.querySelectorAll('h3.ms-linksection-title').forEach(function(el) {
                    if (el.innerHTML.indexOf("Users and Permissions") > 0) {
                        match = el.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode;
                    }
                })
                var group = match.parentNode.children.querySelector('ul');

                //append new child link
                var li = document.createElement("li")
                li.innerHTML = '<a title="Change site owner." href="/_layouts/15/setrqacc.aspx?type=web">Change Site Owner</a>'
                group.appendChild(li);
            }
        }
    };

    //wait until document ready
    function ready(fn) {
        if (document.readyState != 'loading') {
            fn();
        } else {
            document.addEventListener('DOMContentLoaded', fn);
        }
    };

    //execute
    ready(main);
})();