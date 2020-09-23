function clearEmail(obj) {
     if (obj.value=='<Type your email here>') obj.value = '' ;
}
// getPageScroll()
// Returns array with x,y page scroll values.
// Code from - quirksmode.org
var c;
var si, so; //intervals

function getPageScroll(){
    var yScroll, xScroll;
    if (self.pageYOffset) {
        yScroll = self.pageYOffset;
        xScroll = self.pageXOffset;
    } else if (document.documentElement && document.documentElement.scrollTop){	 // Explorer 6 Strict
        yScroll = document.documentElement.scrollTop;
        xScroll = document.documentElement.scrollLeft;
    } else if (document.body) {// all other Explorers
        yScroll = document.body.scrollTop;
        xScroll = document.body.scrollLeft;
    }
    arrayPageScroll = new Array('',yScroll,xScroll) 
    return arrayPageScroll;
}

// getPageSize()
// Returns array with page width, height and window width, height
// Code from - quirksmode.org

function getPageSize(){	
    var xScroll, yScroll;	

    if (window.innerHeight && window.scrollMaxY) {	
        xScroll = document.body.scrollWidth;
        yScroll = window.innerHeight + window.scrollMaxY;
    } else if (document.body.scrollHeight > document.body.offsetHeight){ // all but Explorer Mac
        xScroll = document.body.scrollWidth;
        yScroll = document.body.scrollHeight;
    } else { // Explorer Mac...would also work in Explorer 6 Strict, Mozilla and Safari
        xScroll = document.body.offsetWidth;
        yScroll = document.body.offsetHeight;
    }

    var windowWidth, windowHeight;
	
    if (self.innerHeight) {	// all except Explorer
        windowWidth = self.innerWidth;
        windowHeight = self.innerHeight;
    } else if (document.documentElement && document.documentElement.clientHeight) { // Explorer 6 Strict Mode
        windowWidth = document.documentElement.clientWidth;
        windowHeight = document.documentElement.clientHeight;
    } else if (document.body) { // other Explorers
        windowWidth = document.body.clientWidth;
        windowHeight = document.body.clientHeight;
    }		

    // for small pages with total height less then height of the viewport

    if(yScroll < windowHeight){
        pageHeight = windowHeight;
    } else { 
        pageHeight = yScroll;
    }

    // for small pages with total width less then width of the viewport

    if(xScroll < windowWidth){	
        pageWidth = windowWidth;
    } else {
        pageWidth = xScroll;
    }

    arrayPageSize = new Array(pageWidth,pageHeight,windowWidth,windowHeight) 
    return arrayPageSize;
}

function getScrollerWidth() {
    var scr = null;
    var inn = null;
    var wNoScroll = 0;
    var wScroll = 0;

    // Outer scrolling div
    scr = document.createElement('div');
    scr.style.position = 'absolute';
    scr.style.top = '-1000px';
    scr.style.left = '-1000px';
    scr.style.width = '100px';
    scr.style.height = '50px';
    // Start with no scrollbar
    scr.style.overflow = 'hidden';

    // Inner content div
    inn = document.createElement('div');
    inn.style.width = '100%';
    inn.style.height = '200px';

    // Put the inner div in the scrolling div
    scr.appendChild(inn);
    // Append the scrolling div to the doc
    document.body.appendChild(scr);

    // Width of the inner div sans scrollbar
    wNoScroll = inn.offsetWidth;
    // Add the scrollbar
    scr.style.overflow = 'auto';
    // Width of the inner div width scrollbar
    wScroll = inn.offsetWidth;

    // Remove the scrolling div from the doc
    document.body.removeChild(
    document.body.lastChild);

    // Pixel width of the scroller
    //return (wNoScroll - wScroll);
return 0


}


function doResize() {
  overlay = document.getElementById("overlay");
  overlay.style.width = '100%';

  var arrayPageSize = getPageSize();

  overlay.style.height = arrayPageSize[1] + 'px';
  overlay.style.width = arrayPageSize[0] - getScrollerWidth() + 'px';
  c.style.left = (arrayPageSize[2]/2)-(c.offsetWidth/2)+"px";
}

function doScroll() {
  var arrayPageScroll = getPageScroll();
  c = document.getElementById('flyin');
  c.style.top =  arrayPageScroll[1] + 200 + 'px';
}

function fadeIn (){

  if (document.all) toggleSelects('hidden');

  var arrayPageSize = getPageSize();

  overlay = document.createElement("div");
  overlay.id = 'overlay';

  setOpacity(overlay,75);
  
  document.body.appendChild(overlay);
  overlay.style.height = arrayPageSize[1] + 'px';
  overlay.style.width = arrayPageSize[0] - getScrollerWidth() + 'px';
}


function fadeOut () {    
  setOpacity(overlay,0);  
  c.style.display = 'none';
  overlay.style.display = 'none';
}


function setOpacity(obj, opacity) {
  opacity = (opacity == 100)?99.999:opacity;  
  obj.style.opacity = opacity/100;
  obj.style.MozOpacity= opacity/100;
  obj.style.KhtmlOpacity= opacity/100;
  obj.style.filter = 'alpha(opacity:'+opacity+')';
}

function toggleSelects(state) {
  selects = document.getElementsByTagName('select');
  for (i=0;i<selects.length;i++) selects[i].style.visibility = state;
}

function startSlide(direction) {

  document.forms[0].onsubmit = function () {return false;} //Stop form from submitting in opera


  var arrayPageSize = getPageSize();
  var arrayPageScroll = getPageScroll();

  c = document.getElementById('flyin');
  document.forms[0].appendChild(c);

  if (direction=='down') {
    fadeIn();    
    c.style.display = 'block';
    c.style.top =  ((c.offsetHeight * -1) + arrayPageScroll[1]) + 'px';
    c.style.left = (arrayPageSize[2]/2)-(c.offsetWidth/2)+"px";
    si = setInterval("slideIn()",25);
    document.forms[0].onsubmit = function () {return false;} //Stop form from submitting in opera
  } else {
    so = setInterval("slideOut()",25);
    document.forms[0].onsubmit = function () {return} 
  }
}

function slideIn(container,direction) {
  var arrayPageScroll = getPageScroll();
  ct = parseInt(c.style.top)
  ct < (arrayPageScroll[1] + 200) ? c.style.top = (ct + 30) + 'px' : clearInterval(si);
}

function slideOut(container) {
  var arrayPageScroll = getPageScroll();
  ct = parseInt(c.style.top);    
  if (ct > (c.offsetHeight * -1) + arrayPageScroll[1]) {
    c.style.top = (ct-30) + 'px'
  } else {
    clearInterval(so);
    toggleSelects('visible');
    fadeOut();
  }
}

function switchElements() {  
  document.getElementById('step1').style.display='none';
  document.getElementById('step2').style.display='block';
}

////////////////
//Cookie Stuff
////////////////
function tagUser(expires,status) {
        if (!expires) expires= 3650;
        if (!readCookie('eg_signup' == 1))
        createCookie('eg_signup',1,expires);
}

function createCookie(name,value,days) {
  if (days) {
    var date = new Date();
    date.setTime(date.getTime()+(days*24*60*60*1000));
    var expires = "; expires="+date.toGMTString();
  }
  else var expires = "";
  document.cookie = name+"="+value+expires+"; path=/";
}

function readCookie(name) {
  var nameEQ = name + "=";
  var ca = document.cookie.split(';');
  for(var i=0;i < ca.length;i++) {
    var c = ca[i];
    while (c.charAt(0)==' ') c = c.substring(1,c.length);
    if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
  }
  return null;
}

function cookiesEnabled() {
    var cookieEnabled=(navigator.cookieEnabled)? true : false;

    //if navigator,cookieEnabled is not supported
    if (typeof navigator.cookieEnabled=="undefined" && !cookieEnabled){ 
        document.cookie="testcookie";
        cookieEnabled=(document.cookie.indexOf("testcookie")!=-1)? true : false;
    }
    return cookieEnabled;
}



///////////////
//Email Stuff
///////////////
function AddEmail(email,name,source) {
  if (validateEmail(email)==false) return false
  /*
  document.getElementById('fields').style.display='none';
  document.getElementById("status").style.display="block";
  document.getElementById("status").innerHTML = '<img src="http://www.koders.com/media/images/koders/spinner.gif"> Please wait, we are attempting to add you to our mailing list...<br><span style="font-size:11px;">Your download should automatically begin, if not, click <a href="http://www.koders.com/corp/products/pro/thank-you/">here</a>. </span>';
  */
  if (readCookie('eg_signup') == 1) 
    OnEmailSucceed();
  else
  Koders.WebServices.MailingListService.AddToMailingList(email, null, source, OnEmailSucceed, OnEmailFail);
}

function OnEmailSucceed() {  
  tagUser();
  startSlide('up');          
}

function OnEmailFail() { 
  /*document.getElementById('fields').style.display='none';
  document.getElementById("status").style.display="block";
  document.getElementById("status").innerHTML = 'We\'re sorry, we could not add you to our mailing list at this moment.';
  return false;
  */
  tagUser();
  startSlide('up');
}

function validateEmail() {

  var emailId=document.getElementById('email_address');
  var emailStatus = document.getElementById('email_status');
  var emailPat = /^([a-zA-Z0-9_\.\-+])+@(([a-zA-Z0-9-])+\.)+([a-zA-Z0-9]{2,4})+$/;

  if(emailId.value.length==0) {
    emailStatus.innerHTML = "Please enter an email address.";
    emailStatus.style.display = 'block';
    emailId.focus()
    return false;
  }

  if (!(emailId.value.match(emailPat))) {
    emailStatus.innerHTML = "Please enter a valid email address.";
    emailStatus.style.display = 'block';
    return false;
  }
  emailStatus.style.display = 'none';
  return true;
}

function checkForSubmit(e) {
   if(window.event) // IE
    {
    keynum = e.keyCode
    }
    else if(e.which) // Netscape/Firefox/Opera
    {
    keynum = e.which
    }
    
    if (keynum == 13) {  
      AddEmail(document.getElementById('email_address').value,null,'eg2');
      return false
    }
    return true
}

function trackUser() {
    var trackerImage = new Image;
    trackerImage.src = 'http://www.koders.com/Special/AdServer/?action=serveimage&flight=85&imageurl=KD7KH2LHS3NR157RR348GVYPV566VU7QQ5N4ZVUPLN28VXYQB53MZ37FN55KZVLRW5NGH2FPDN3EH16PLW3A&random=633234786820607376';
    document.getElementById("flyin").appendChild(trackerImage);                    
}        

if (cookiesEnabled() && readCookie('eg_signup') != 1) {
        setTimeout('startSlide("down")',1000);
        window.onload = trackUser;
        window.onresize = doResize;
        window.onscroll = doScroll;
}