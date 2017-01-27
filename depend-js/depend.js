function depend(src, callback) {
	// Already loaded
	var tags = document.getElementsByTagName('script')
	for (var i = 0; i < tags.length; i++) {
		if (tags[i].src == src) {
			console.log('dep-found');
			callback();
			return;
		}
	}

	// First load
	console.log('dep-new ' + src);
	var tag = document.createElement('script');
	tag.src = src;
	tag.onload = callback;
	document.head.appendChild(tag); //or something of the likes
}