//referenced code at http://stackoverflow.com/questions/16149431/make-function-wait-until-element-exists

function waitDom(id, fn) {
    var checkExist = setInterval(function() {
        if (document.getElementById(id)) {
            console.log("OK!");
            clearInterval(checkExist);
            fn();
        } else {
            console.log('wait...');
        }
    }, 500);
}

function hello() {
    alert('hello');
}

function addContainer() {
    document.write("<div id='container'></div>");
}

//main
waitDom('container', hello);
window.setTimeout(addContainer, 2000);