import { getParameter } from './script.js';

var user = getParameter('user', '');
var url = getParameter('url', '');
document.getElementById('user').innerHTML = user;
if (url !== "") {
    var title = 'eMailerOOo'
    if (user !== "") {
        title = title + ' - ' + user
    }
    window.open(url, title, 'width=800,height=600');
    window.close();
}