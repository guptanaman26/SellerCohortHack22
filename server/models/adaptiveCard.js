const { StatusCodes } = require('botbuilder');
var fs = require("fs");
var axios = require("axios");
var favouriteController = require('../cardLoadController/favouriteController');
var accountSummaryCard=require ('./accountSummaryCard');
var tpidData = require ('./sellerAccountMapping');
const { response } = require('express');
// Card response for authentication
const createAuthResponse = (signInLink) => {
    console.log("Create Auth response")
    const res = {
            tab: {
                type: "auth",
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: signInLink,
                            title: "Sign in to this app"
                        }
                    ]
                }
            }
    };
    return res;
};

// Card response for task module invoke request
const invokeTaskResponse = () => {
    const response = {
            task: {
                type: 'continue',
                value: {
                    card: {
                        contentType: "application/vnd.microsoft.card.adaptive",
                        content: adaptiveCardTaskModule
                    },
                    heigth: 250,
                    width: 400,
                    title: 'Sample Adaptive Card'
                }
            }
    };
    return response;
};

const videoInvokeResponse = (videoId) => {
    const response = {
            task: {
                type: 'continue',
                value: {
                    url: `https://www.youtube.com/embed/${videoId}`,
                    fallbackUrl: `https://www.youtube.com/embed/${videoId}`,
                    heigth: 1000,
                    width: 700,
                    title: 'Youtube Video'
                }
            }
    };
    return response;
};

// Card response for tab fetch request
const createFetchResponse = async (userImage, displayName, userId) => {
    console.log("Create Invoke response");
    var imageString = '';
    if (userImage) {
        // Converting image of Blob type to base64 string for rendering as image.
        await userImage.arrayBuffer().then(result => {
            console.log(userImage.type);
            imageString = Buffer.from(result).toString('base64');
            if (imageString != '') {
                // Writing file to Images folder to use as url in adaptive card
                fs.writeFileSync("Images/profile-image.jpeg", imageString, { encoding: 'base64' }, function (err) {
                    console.log("File Created");
                });
            }
        }).catch(error => { console.log(error) });
    }

    var cardsArray = [
        {
            "card": getAdaptiveCardUserDetails(imageString, displayName)
        }
    ];

    var favouriteAccount = await fetchFavoruiteAccount(userId);
    favouriteAccount.forEach(eachAccount => {
        cardsArray.push({"card":accountSummaryCard.getAccountSummaryCard(eachAccount,userId)});
    });

    cardsArray.push({"card": getSearchButtonCard(userId)})


    const res = {
            tab: {
                type: "continue",
                value: {
                    cards: cardsArray
                },
        }
    };

    return res;
};

const getSearchButtonCard=(userId)=>{
    var response ={
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "horizontalAlignment": "Center",
                                "verticalContentAlignment": "Center",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "Find More Accounts",
                                                "data":{
                                                    "loadView":"initialSearch",
                                                    "userId":userId
                                                }
                                            }
                                        ]
                                    }
                                ]
                            }
                        ],
                        "horizontalAlignment": "Center"
                    }
                ],
                "horizontalAlignment": "Center"
            }
        ]
    };
    return response;
}

function getAccountName(tpid) {
    for (var i = 0; i < tpidData.TpidAccountMapping.length; i++) {
      if (tpidData.TpidAccountMapping[i]["TPID"] === tpid ) return tpidData.TpidAccountMapping[i]["AccountName"];
    }
    return null;
  }
  const getDecisionMakerCount = async (tpid) => {
    try {
        const authToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJmYTUwMjIwNy1iOWM1LTQ0MDAtYmIxYy1mYTIzYzIwNDIyYWYiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDcvIiwiaWF0IjoxNjY0MTEyNTY1LCJuYmYiOjE2NjQxMTI1NjUsImV4cCI6MTY2NDE5OTI2NSwiYWlvIjoiRTJaZ1lQaXRPTlBvbGM3am45Y20rYTBRODhrOURnQT0iLCJhcHBpZCI6IjgwZGY5OWY2LWJmYzktNGM0Ny1iNzI2LTQyYjY3Y2NlMzUwMyIsImFwcGlkYWNyIjoiMSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0Ny8iLCJvaWQiOiI0OGFjZDY4Ny1jYmIzLTQ5NzAtYWRhNC02MzdlZWM1MGI3ZWYiLCJyaCI6IjAuQVJvQXY0ajVjdkdHcjBHUnF5MTgwQkhiUndjaVVQckZ1UUJFdXh6Nkk4SUVJcThhQUFBLiIsInN1YiI6IjQ4YWNkNjg3LWNiYjMtNDk3MC1hZGE0LTYzN2VlYzUwYjdlZiIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsInV0aSI6IklaVmMzUFJmMUVHbUVmbWNSaEVSQUEiLCJ2ZXIiOiIxLjAifQ.b6FpK-d6UZNmQqoRI0AP59vkY9EZF0Q9ZfQCILhD2NILtqvTKekZQHkrI7gS-1SNuVUqtppCPLSBX3Yj5QELVciSWXgJUdCfSP7-3FUStvyb_NFdhdc7IXuAePt5-z6r1V-RXDPXkTQElfFzDXGs3-D4oxCrq8ac9S39zZbApjEnsmEFddBbf_qKO5wUzQEPN4v25tl3_-nN4rKXDIs_YQQ4jWgxEyUGbIjqQQyJ74bUtpqF76XMBQI4pxI2_DQ661E4uh-PwQYFyHNHRU0HKpKXoZbUAGNnovyRN2FWBN0t3Ylh9XrKFLVJ1vx84S0jurnI63Bt99LT6SFLS9QqSg";
        const config = {
            headers: { 
            Authorization: `Bearer ${authToken}`,
            CorrelationId: "bea4d80e-7363-435d-94d4-7635522a896c",
            pageSize: 50,
            pageNumber: 1
            }
      };
      let url= `https://activitystore-ppe.trafficmanager.net/composite/api/v1/Activity/Details/CSP?Key=TPID&value=${tpid}`;
      let resp = await axios.get(url, config)
      let actualData = await resp.data;
      let result = 0;
      actualData["activities"].forEach(e => {
        if(e.activityDetail.marketingAudience &&
            (e.activityDetail.marketingAudience.includes("BDM") 
            || e.activityDetail.marketingAudience.includes("ITDM"))
          ) 
          result++;
      });
      return result;
    } catch(err) {
      console.log(err.message)
        return (err.message);
    }
  }

    const getAccountSummaryList = async (tpidArray,userId) => {
      try {
        let result = [];
        if(tpidArray.length==0){
            return result;
        }
        const authToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJmYTUwMjIwNy1iOWM1LTQ0MDAtYmIxYy1mYTIzYzIwNDIyYWYiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDcvIiwiaWF0IjoxNjY0MTEyNTY1LCJuYmYiOjE2NjQxMTI1NjUsImV4cCI6MTY2NDE5OTI2NSwiYWlvIjoiRTJaZ1lQaXRPTlBvbGM3am45Y20rYTBRODhrOURnQT0iLCJhcHBpZCI6IjgwZGY5OWY2LWJmYzktNGM0Ny1iNzI2LTQyYjY3Y2NlMzUwMyIsImFwcGlkYWNyIjoiMSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0Ny8iLCJvaWQiOiI0OGFjZDY4Ny1jYmIzLTQ5NzAtYWRhNC02MzdlZWM1MGI3ZWYiLCJyaCI6IjAuQVJvQXY0ajVjdkdHcjBHUnF5MTgwQkhiUndjaVVQckZ1UUJFdXh6Nkk4SUVJcThhQUFBLiIsInN1YiI6IjQ4YWNkNjg3LWNiYjMtNDk3MC1hZGE0LTYzN2VlYzUwYjdlZiIsInRpZCI6IjcyZjk4OGJmLTg2ZjEtNDFhZi05MWFiLTJkN2NkMDExZGI0NyIsInV0aSI6IklaVmMzUFJmMUVHbUVmbWNSaEVSQUEiLCJ2ZXIiOiIxLjAifQ.b6FpK-d6UZNmQqoRI0AP59vkY9EZF0Q9ZfQCILhD2NILtqvTKekZQHkrI7gS-1SNuVUqtppCPLSBX3Yj5QELVciSWXgJUdCfSP7-3FUStvyb_NFdhdc7IXuAePt5-z6r1V-RXDPXkTQElfFzDXGs3-D4oxCrq8ac9S39zZbApjEnsmEFddBbf_qKO5wUzQEPN4v25tl3_-nN4rKXDIs_YQQ4jWgxEyUGbIjqQQyJ74bUtpqF76XMBQI4pxI2_DQ661E4uh-PwQYFyHNHRU0HKpKXoZbUAGNnovyRN2FWBN0t3Ylh9XrKFLVJ1vx84S0jurnI63Bt99LT6SFLS9QqSg";
          const config = {
              headers: { 
              Authorization: `Bearer ${authToken}`,
              CorrelationId: "bea4d80e-7363-435d-94d4-7635522a896c",
              pageSize: 50,
              pageNumber: 1
              }
        };
        let url= `https://activitystore-ppe.trafficmanager.net/composite/api/v1/Activity/Summary?key=TPID&value=${tpidArray.join(",")}`;
        let resp = await axios.get(url, config);
        let actualData = await resp.data;
        for(var i = 0; i<actualData["activitySummary"]["details"].length; i++){
            var e = actualData["activitySummary"]["details"][i];
            let element = {
                    "accountName": getAccountName(e.tpid),
                    "tpId":e.tpid,
                    "favourite": favouriteController.isFavourite(e.tpid, userId),
                    "summaryData":
                    {
                      "totalContacts": e.ccmCount,
                      "totalInteractions": e.csp.totalInteractions,
                      "interactionsByDM": await getDecisionMakerCount(e.tpid),
                      "interactionGain": (Math.random() < 0.7 ? 1 : -1) * (Math.floor(Math.random() * (600-100) + 100)/100)
                    }
                  }
                  result.push(element);
        }
        console.log(result);
        return result;
      } catch(err) {
        console.log(err.message)
          return (err.message);
      }
    }

const getFavoriteTpidArray=(userId)=>{
    let response=[];
    tpidData.TpidAccountMapping.forEach(element=>{
        if(element.SellerID==userId && favouriteController.isFavourite(element.TPID,userId)){
            response.push(element.TPID);
        }
    });
    return response;
}

const fetchFavoruiteAccount = async (userId) =>{
    const favouriteTpidArray = getFavoriteTpidArray(userId);
    const res = await getAccountSummaryList(favouriteTpidArray,userId);
    return res;
}

// Card response for tab fetch request
const createFetchResponseForTab2 = async () => {
    console.log("Create Invoke response");

    const res = {
            tab: {
                type: "continue",
                value: {
                    cards: [
                        {
                            "card": getAdaptiveCardTab2(),
                        }
                    ]
                },
            }
    };

    return res;
};

// Card response for tab submit request
const createSubmitResponse = () => {
    console.log("Submit response")
    const res = {
            tab: {
                type: "continue",
                value: {
                    cards: [
                        {
                            card: signOutCard,
                        }
                    ]
                },
            }
    };

    return res;
};

// Card response for tab submit request
const taskSubmitResponse = () => {
    console.log("Task Submit response")
    const response = {
        task: {
            type: 'continue',
            value: {
                card: {
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content: taskSubmitCard
                },
                heigth: 250,
                width: 400,
                title: 'Sample Adaptive Card'
            }
        }
};
return response;
};

// Adaptive Card with user image, name and Task Module invoke action
const getAdaptiveCardTab2 = () => {
    const adaptiveCard1 = {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        body: [
            {
                type: "Container",
                items: [
                  {
                    type: "TextBlock",
                    text: "Video Player",
                    weight: "Bolder",
                    size: "Medium"
                  }
                ]
              },
              {
                type: "Container",
                items: [
                  {
                    type: "TextBlock",
                    text: "Enter the ID of a YouTube video to play in a dialog",
                    wrap: true
                  },
                  {
                    type: "Input.Text",
                    id: "youTubeVideoId",
                    value: "jugBQqE_2sM"
                  }
                ]
              },
            {
                type: 'ActionSet',
                actions: [
                    {
                        type: "Action.Submit",
                        title: "Show Task Module",
                        data: {
                            msteams: {
                                type: "task/fetch"
                            }
                        }
                    }
                ]
            }
        ],
        type: 'AdaptiveCard',
        version: '1.4'
    };
    return adaptiveCard1;
}

// Adaptive Card with user image, name and Task Module invoke action
const getAdaptiveCardUserDetails = (image, name) => {
    const adaptiveCard1 = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": "https://messagecardplayground.azurewebsites.net/assets/Mostly%20Cloudy-Square.png",
                                        "size": "Small",
                                        "altText": "UserProfileImage",
                                        "width": "50px",
                                        "horizontalAlignment": "Right"
                                    }
                                ],
                                "verticalContentAlignment": "Center"
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Hello "+ name + ",",
                                        "wrap": true,
                                        "size": "Large",
                                        "weight": "Bolder",
                                        "color": "Accent"
                                    }
                                ],
                                "horizontalAlignment": "Left",
                                "verticalContentAlignment": "Center"
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Input.ChoiceSet",
                                        "id":"homeView",
                                        "choices": [
                                            {
                                                "title": "Favourite Acccounts",
                                                "value": "favourite"
                                            },
                                            {
                                                "title": "Interaction Gain ",
                                                "value": "interactionGain"
                                            }
                                        ],
                                        "placeholder": "Select View",
                                        "value": "favourite"
                                    },
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "View",
                                                "style": "positive"
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }
        ]
    };
    return adaptiveCard1;
}

// Adaptive Card showing sample text and Submit Action
const getAdaptiveCardSubmitAction = () => {
    const adaptiveCard2 = {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        body: [
            {
                type: 'Image',
                height: '300px',
                width: '400px',
                url: 'https://cdn.vox-cdn.com/thumbor/Ndb49Uk3hjiquS041NDD0tPDPAs=/0x169:1423x914/fit-in/1200x630/cdn.vox-cdn.com/uploads/chorus_asset/file/7342855/microsoftteams.0.jpg',
            },
            {
                type: 'TextBlock',
                size: 'Medium',
                weight: 'Bolder',
                text: 'tab/fetch is the first invoke request that your bot receives when a user opens an Adaptive Card tab. When your bot receives the request, it either sends a tab continue response or a tab auth response',
                wrap: true,
            },
            {
                type: 'TextBlock',
                size: 'Medium',
                weight: 'Bolder',
                text: 'tab/submit request is triggered to your bot with the corresponding data through the Action.Submit function of Adaptive Card',
                wrap: true,
            },
            {
                type: 'ActionSet',
                actions: [
                    {
                        type: 'Action.Submit',
                        title: 'Sign Out',
                    }
                ],
            }
        ],
        type: 'AdaptiveCard',
        version: '1.4'
    };
    return adaptiveCard2;
}

// Adaptive Card to show in task module
const adaptiveCardTaskModule = {
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    body: [
        {
            type: 'TextBlock',
            size: 'Medium',
            weight: 'Bolder',
            text: 'Sample task module flow for tab'
        },
        {
            type: 'Image',
            height: '50px',
            width: '50px',
            url: 'https://cdn.vox-cdn.com/thumbor/Ndb49Uk3hjiquS041NDD0tPDPAs=/0x169:1423x914/fit-in/1200x630/cdn.vox-cdn.com/uploads/chorus_asset/file/7342855/microsoftteams.0.jpg',
        },
        {
            type: 'ActionSet',
            actions: [
                {
                    type: "Action.Submit",
                    title: "Close",
                    data: {
                        msteams: {
                            type: "task/submit"
                        }
                    }
                }
            ]
        }
    ],
    type: 'AdaptiveCard',
    version: '1.4'
};

// Adaptive Card to show sign out action
const signOutCard = {
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    body: [
        {
            type: 'TextBlock',
            size: 'Medium',
            weight: 'Bolder',
            text: 'Sign out successful. Please refresh to Sign in again.',
            wrap: true,
        }
    ],
    type: 'AdaptiveCard',
    version: '1.4'
};

// Adaptive Card to show task/submit action
const taskSubmitCard = {
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    body: [
        {
            type: 'TextBlock',
            size: 'Medium',
            weight: 'Bolder',
            text: 'The action called task/submit. Please refresh to laod contents again.',
            wrap: true,
        }
    ],
    type: 'AdaptiveCard',
    version: '1.4'
};

module.exports = {
    createFetchResponse,
    createFetchResponseForTab2,
    createSubmitResponse,
    createAuthResponse,
    invokeTaskResponse,
    videoInvokeResponse,
    taskSubmitResponse,
    getAccountSummaryList
};