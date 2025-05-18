
// document.getElementById("userInputButton").addEventListener("click", getUserInput, false);
// document.getElementById("userInput").addEventListener("keyup", function (event) {...});

eel.expose(addUserMsg);
eel.expose(addAppMsg);

function addUserMsg(msg) {
    let element = document.getElementById("messages");
    element.innerHTML += '<div class="message from ready rtol">' + msg + '</div>';
    element.scrollTop = element.scrollHeight - element.clientHeight - 15;
    let index = element.childElementCount - 1;
    setTimeout(changeClass.bind(null, element, index, "message from"), 500);
}

function addAppMsg(msg) {
    let element = document.getElementById("messages");
    element.innerHTML += '<div class="message to ready ltor">' + msg + '</div>';
    element.scrollTop = element.scrollHeight - element.clientHeight - 15;
    let index = element.childElementCount - 1;
    setTimeout(changeClass.bind(null, element, index, "message to"), 500);
}

function changeClass(element, index, newClass) {
    console.log(newClass + ' ' + index);
    element.children[index].className = newClass;
}
