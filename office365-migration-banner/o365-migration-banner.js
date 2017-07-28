// O365 - Migration Banner
function addcss(css) {
	var head = document.getElementsByTagName('head')[0];
	var s = document.createElement('style');
	s.setAttribute('type', 'text/css');
	if (s.styleSheet) {
		// IE
		s.styleSheet.cssText = css;
	} else {
		// the world
		s.appendChild(document.createTextNode(css));
	}
	head.appendChild(s);
}

// CSS Inject
addcss('#status_preview_body {display:none}');

// Migration Function
MigrationBannerCount = 0;
function MigrationBanner() {
	var sp = document.getElementById('status_preview');
	if (MigrationBannerCount > 10) {
		console.log('MigrationBanner - safety, max attempts, to prevent infinite loop');
		//safety, max attempts, to prevent infinite loop
		return
	}
	if (!sp) {
		console.log('MigrationBanner - wait and check later');
		//wait and check later
		MigrationBannerCount++;
		window.setTimeout(MigrationBanner, 200);
		return;
	} else {
		console.log('MigrationBanner - found and modify');
		//found and modify
		var h = sp.innerHTML;
		h = h.replace("This site is read only at the farm administrator's request.", '<table border=0><tr><td><img src="/_layouts/images/kpinormallarge-1.gif" height="15px" width="15px" style="padding-right:20px"/></td><td><b>MIGRATION IN PROGRESS:</b> This site is being moved to Office 365 and will be locked as Read-Only until its migration is complete.<br/>Please check in soon to be automatically redirected to this site’s new location in the cloud. <a href="https://kbarticle">Learn more...</a></td></tr><table>');
		sp.innerHTML = h;
		addcss('#status_preview_body {display:inherit}');
	}
}
function MigrationBannerLoad() {
	ExecuteOrDelayUntilScriptLoaded(MigrationBanner, "sp.js")
}
// Delay Load
window.setTimeout('MigrationBannerLoad()', 1000);