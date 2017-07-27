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
		sp.innerHTML = h;
		addcss('#status_preview_body {display:inherit}');
	}
}
function MigrationBannerLoad() {
	ExecuteOrDelayUntilScriptLoaded(MigrationBanner, "sp.js")
}
// Delay Load
window.setTimeout('MigrationBannerLoad()', 1000);
