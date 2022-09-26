const { StatusCodes } = require('botbuilder');
var fs = require("fs");
var accountDetailsModule = require('../models/accountDetailsCard');
 
const fetchAccountDetailsCard =async (data)=>{
    var cardsArray = [
        {
            card: accountDetailsModule.getFilterCard(data),
        },
        {
            card: await accountDetailsModule.getInteractedDetails(data),
        }
    ]
    if(data.salesPlay!=undefined && data.salesPlay != ''){
        cardsArray.push({
            card: {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [
                    {
                        "type": "Container",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Interacted Contacts",
                                "wrap": true,
                                "size": "Large",
                                "weight": "Bolder"
                            }
                        ],
                        "style": "emphasis"
                    }
                ]
            }
        });
        var interactionData = await accountDetailsModule.fetchDecisionMakerData(data?.tpId, data.solutionArea, data.salesPlay);
        interactionData.forEach(element => {
            cardsArray.push({
                card: accountDetailsModule.getDecisionMakerCard(element)
            });
        });
    }
    const res = {
            tab: {
                type: "continue",
                value: {
                    cards: cardsArray
                },
            }
    };
    console.log(cardsArray);
    return res;
}

module.exports = {
    fetchAccountDetailsCard
};