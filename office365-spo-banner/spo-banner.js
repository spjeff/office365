// Inject CSS
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

// Set cookie
function setCookie(cname, cvalue, exdays) {
  var d = new Date();
  d.setTime(d.getTime() + exdays * 24 * 60 * 60 * 1000);
  var expires = 'expires=' + d.toUTCString();
  document.cookie = cname + '=' + cvalue + ';' + expires + ';path=/';
}

// Read cookie
function getCookie(cname) {
  var name = cname + '=';
  var decodedCookie = decodeURIComponent(document.cookie);
  var ca = decodedCookie.split(';');
  for (var i = 0; i < ca.length; i++) {
    var c = ca[i];
    while (c.charAt(0) == ' ') {
      c = c.substring(1);
    }
    if (c.indexOf(name) == 0) {
      return c.substring(name.length, c.length);
    }
  }
  return '';
}

// Display banner
function spoBannerLoad() {
  // HTTP GET for banner SPList items
  var request = new XMLHttpRequest();
  request.open('GET', "/_api/web/lists/getbytitle('GlobalAlert')/items", true);
  request.setRequestHeader('Accept', 'application/json; odata=verbose');

  request.onload = function() {
    if (request.status >= 200 && request.status < 400) {
      // HTTP 200 - Success!
      var data = JSON.parse(request.responseText);
      var title = data.d.results[0].Title;
      var id = data.d.results[0].Id;

      // Read cookie
      var cookie = getCookie('spo-banner');
      if (cookie.indexOf(id) > -1) {
        // Not found
        return;
      }

      // Inject HTML <div>
      var body = document.getElementsByTagName('body')[0];
      var d = document.createElement('div');
      d.id = 'spo-banner';
      d.innerHTML =
        '<span><b>' +
        title +
        '</b><sup><font size="2em" style="padding-left:10px"><a href="javascript:void(0)" onclick="spoBannerDismiss(' +
        id +
        ')">[X]</font></a></sup></span>';
      body.appendChild(d);

      // Calculate left offset
      var d = document.getElementById('spo-banner');
      var left = window.innerWidth / 2 - 100;

      // Inject CSS
      addcss(
        '#spo-banner {background-color: yellow;padding:3px;top:0px;left:' +
          left +
          'px;z-index: 99;position: fixed;}'
      );
    } else {
      // We reached our target server, but it returned an error
    }
  };
  request.onerror = function() {
    // There was a connection error of some sort
  };

  request.send();
}

// Dismiss banner
function spoBannerDismiss(id) {
  // User click [X] to dismiss. Set cookie.
  var d = document.getElementById('spo-banner');
  d.style.display = 'none';
  setCookie('spo-banner', id, 7);
}

// Main
spoBannerLoad();
