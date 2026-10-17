const english = navigator.language.toLowerCase().startsWith('en');
document.documentElement.lang = english ? 'en' : 'zh-TW';
document.querySelectorAll('[data-zh]').forEach(element => {
    element.textContent = element.dataset[english ? 'en' : 'zh'];
});
function show(enabled) {
    document.body.classList.toggle('state-on', enabled === true);
    document.body.classList.toggle('state-off', enabled === false);
}
document.querySelector('.open-preferences').addEventListener('click', () => {
    webkit.messageHandlers.controller.postMessage('open-preferences');
});
document.querySelector('.open-e3').addEventListener('click', () => {
    webkit.messageHandlers.controller.postMessage('open-e3');
});
