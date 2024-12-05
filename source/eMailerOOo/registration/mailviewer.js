import { getParameter } from './script.js';

var user = getParameter('user', '');
document.getElementById('user').innerHTML = user;
var url = getParameter('url', '');
var title = 'eMailerOOo - ' + user
window.open(url, title, 'width=800,height=600');
window.close();