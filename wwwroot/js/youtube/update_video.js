//alert('upload_video.js script loaded');
/***** START BOILERPLATE CODE: Load client library, authorize user. *****/

  // Global variables for GoogleAuth object, auth status.
  var GoogleAuth;

  /**
   * Load the API's client and auth2 modules.
   * Call the initClient function after the modules load.
   */
  function handleClientLoad() {
    gapi.load('client:auth2', initClient);
  }

  function initClient() {
    // Initialize the gapi.client object, which app uses to make API requests.
    // Get API key and client ID from API Console.
    // 'scope' field specifies space-delimited list of access scopes

    gapi.client.init({
        'clientId': '153908722054-9d66tajejhdnlceondmvk790iaobun42.apps.googleusercontent.com',
        'discoveryDocs': ['https://www.googleapis.com/discovery/v1/apis/youtube/v3/rest'],
        'scope': 'https://www.googleapis.com/auth/youtube.force-ssl https://www.googleapis.com/auth/youtubepartner'
    }).then(function () {
      GoogleAuth = gapi.auth2.getAuthInstance();

      // Listen for sign-in state changes.
      GoogleAuth.isSignedIn.listen(updateSigninStatus);

      // Handle initial sign-in state. (Determine if user is already signed in.)
      setSigninStatus();

      // Call handleAuthClick function when user clicks on "Authorize" button.
      $('#execute-request-button').click(function() {
        handleAuthClick(event);
      }); 
    });
  }

  function handleAuthClick(event) {
    // Sign user in after click on auth button.
    GoogleAuth.signIn();
  }

  var allowRequest = false;  

  function setSigninStatus() {
    var user = GoogleAuth.currentUser.get();
    isAuthorized = user.hasGrantedScopes('https://www.googleapis.com/auth/youtube.force-ssl https://www.googleapis.com/auth/youtubepartner');
    // Toggle button text and displayed statement based on current auth status.
    if (isAuthorized) {
      //defineRequest();
        allowRequest = true;
    }
  }

  function updateSigninStatus(isSignedIn) {
    setSigninStatus();
  }

  function createResource(properties) {
    var resource = {};
    var normalizedProps = properties;
    for (var p in properties) {
      var value = properties[p];
      if (p && p.substr(-2, 2) == '[]') {
        var adjustedName = p.replace('[]', '');
        if (value) {
          normalizedProps[adjustedName] = value.split(',');
        }
        delete normalizedProps[p];
      }
    }
    for (var p in normalizedProps) {
      // Leave properties that don't have values out of inserted resource.
      if (normalizedProps.hasOwnProperty(p) && normalizedProps[p]) {
        var propArray = p.split('.');
        var ref = resource;
        for (var pa = 0; pa < propArray.length; pa++) {
          var key = propArray[pa];
          if (pa == propArray.length - 1) {
            ref[key] = normalizedProps[p];
          } else {
            ref = ref[key] = ref[key] || {};
          }
        }
      };
    }
    return resource;
  }

  function removeEmptyParams(params) {
    for (var p in params) {
      if (!params[p] || params[p] == 'undefined') {
        delete params[p];
      }
    }
    return params;
  }

  function executeRequest(request) {
    request.execute(function(response) {
      console.log(response);
      console.log("execute request...")
    });
  }

  function buildApiRequest(requestMethod, path, params, properties) {
    params = removeEmptyParams(params);
    var request;
    if (properties) {
      var resource = createResource(properties);
      request = gapi.client.request({
          'body': resource,
          'method': requestMethod,
          'path': path,
          'params': params
      });
    } else {
      request = gapi.client.request({
          'method': requestMethod,
          'path': path,
          'params': params
      });
    }
    executeRequest(request);
  }

  /***** END BOILERPLATE CODE *****/

/**
*   get my strings for youtube title and description from existing database input
*   fields, then use html onblur function to record any changes 
*/

var titleText;

var descriptionText;

var diveDate;

var vesselName;

var diveName;

var diveLocation;

var lat;

var lng;

var myVideoId;

//read existing inputs on document ready:
$(document).ready(function() {
    getDiveDate();
    getVesselName();
    getDiveName();
    getLocation();
    getLatitude();
    getLongitude();

    myVideoId = document.getElementById("MyVideoId").value;
});

//onblur getters records any changes to input fields:

function getDiveDate(){
    diveDate = document.getElementById("DiveDate").value;
}

function getVesselName(){
    vesselName = document.getElementById("VesselName").value;
}

function getDiveName(){
    diveName = document.getElementById("DiveName").value;
}

function getLocation(){
    diveLocation = document.getElementById("Location").value; 
}

function getLatitude(){
    lat = document.getElementById("Latitude").value; 
}

function getLongitude(){
    lng = document.getElementById("Longitude").value; 
}

function setTitleAndDescriptionText(){

    titleText = "ROV Holland I: " + diveName;

    descriptionText = "The Marine Institute operates the Remotely Operated Vehicle (ROV) Holland I from its own research vessels" 
        + "and from vessels operated by other institutions. The Holland I is equipped with a wide range of scientific research devices,"
        + "including a high definition camera system.\n"  
        + "This footage was captured during a " + diveName + " mission on " + diveDate + ".  " 
        + "The ROV was launched from " + vesselName + " in the " + diveLocation + " (lat " + lat + ", " + "lng " + lng + ").";  

console.log("Set: "+titleText);
console.log("Set: "+descriptionText);
}

//assign this to Save button

$( "#execute-update" ).click(function() {
  console.log( "Do button called..." );

    if(allowRequest){

        setTitleAndDescriptionText();//setting the title text and description before update
        console.log("defineRequest: "+titleText);
        console.log("defineRequest: "+descriptionText);

    // See full sample for buildApiRequest() code, which is not 
    // specific to a particular API or API method
     buildApiRequest('PUT',
                '/youtube/v3/videos',
                {'part': 'snippet,status'},
                {'id': myVideoId,
                 'snippet.categoryId': '22',
                 //'snippet.defaultLanguage': '',
                 'snippet.description': descriptionText,
                 //'snippet.tags[]': '',
                 'snippet.title': titleText
                 //'status.privacyStatus': ''
      });

     alert("Youtube updated");      
}else{
     alert("Youtube update unsuccessful. Check authorisation.");
}


});

/*

function defineRequest() {



}
*/

