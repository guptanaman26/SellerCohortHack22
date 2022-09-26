// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, TurnContext } = require('botbuilder');
const adaptiveCards = require('../models/adaptiveCard');
const accountDetailsController = require('../cardLoadController/accountDetailsController');
const accountSearchController = require('../cardLoadController/accountSearchController');
var favouriteController = require('../cardLoadController/favouriteController');
const { GraphClient } = require('../graphClient')
class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
    }

    async handleTeamsTabFetch(context, tabRequest) {
        // // When the Bot Service Auth flow completes, context will contain a magic code used for verification.
        // const magicCode =
        //     context.activity.value && context.activity.value.state
        //         ? context.activity.value.state
        //         : '';
        
        // // Getting the tokenResponse for the user
        // const tokenResponse = await context.adapter.getUserToken(
        //     context,
        //     process.env.connectionName,
        //     magicCode
        // );

        // if (!tokenResponse || !tokenResponse.token) {
        //     // Token is not available, hence we need to send back the auth response

        //     // Retrieve the OAuth Sign in Link.
        //     const signInLink = await context.adapter.getSignInLink(
        //         context,
        //         process.env.ConnectionName
        //     );

        //     // Generating and returning auth response.
        //     return adaptiveCards.createAuthResponse(signInLink);
        // }

        //     const graphClient = new GraphClient(tokenResponse.token);

        //     const profile = await graphClient.GetUserProfile().catch(error => {
        //         console.log(error);
        //     });

        //     const userImage = await graphClient.GetUserPhoto().catch(error => {
        //         console.log(error);
        //     });

            var profile = {
                displayName: "Naman",
                userId: "1f192781-48ba-e511-80e4-3863bb2eb058"
            };
            var userImage = null;

            return await adaptiveCards.createFetchResponse(userImage, profile.displayName, profile.userId);
    }

    async handleTeamsTabSubmit(context, tabRequest) {
        console.log('Trying to submit tab content');

        if(tabRequest.data.loadView=="accountDetails"){
            return await accountDetailsController.fetchAccountDetailsCard(tabRequest.data);
        }
        else if(tabRequest.data.loadView=="initialSearch"){
            tabRequest.data.tpId='';
            tabRequest.data.solutionArea='';
            tabRequest.data.salesPlay='';
            return accountSearchController.fetchInitialAccountSeacrchCards(tabRequest.data);
        }
        else if(tabRequest.data.loadView=="searchResults"){
            return await accountSearchController.fetchAccountSearchCards(tabRequest.data);
        }
        else if(tabRequest.data.loadView=="toggleFavourite"){
            favouriteController.toggleFavourite(tabRequest.data);
        }else if(tabRequest.data.loadView=="clearAccountDetailsFilter"){
            let data = tabRequest.data;
            data.solutionArea = '';
            data.salesPlay = '';
            return await accountDetailsController.fetchAccountDetailsCard(tabRequest.data);
        }
        var profile = {
            displayName: "Naman",
            userId: "1f192781-48ba-e511-80e4-3863bb2eb058"
        };
        var userImage = null;
        return await adaptiveCards.createFetchResponse(userImage, profile.displayName, profile.userId);
    }

    handleTeamsTaskModuleFetch(context, taskModuleRequest) {
        var tabEntityId = taskModuleRequest.tabContext.tabEntityId;
        if (tabEntityId == "youtubeTab") {
            var videoId = taskModuleRequest.data.youTubeVideoId;
            return adaptiveCards.videoInvokeResponse(videoId);
        }
        else
            return adaptiveCards.invokeTaskResponse();
    }

    handleTeamsTaskModuleSubmit(context, taskModuleRequest) {
        return adaptiveCards.taskSubmitResponse();
    }
}

module.exports.BotActivityHandler = BotActivityHandler;