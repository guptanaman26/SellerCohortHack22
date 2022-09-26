const { StatusCodes } = require('botbuilder');
var fs = require("fs");
var accountSearchModule=require('../models/accountSearchCard');
var accountSummaryModule=require('../models/accountSummaryCard');
 
const fetchInitialAccountSeacrchCards =(data)=>{
    var cardsArray = [
        {
            card: accountSearchModule.searchCard(data),
        }
    ];
    const res = {
            tab: {
                type: "continue",
                value: {
                    cards: cardsArray
                },
            }
    };
    return res;
}

const invalidFilterAppliedCard=()=>{
    const card ={
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Please try searching again, No Data Found. Search currently supports only one filter from the above options.",
                        "wrap": true,
                        "weight": "Bolder",
                        "color": "Warning"
                    }
                ]
            }
        ]
    };
    return card;
}

const fetchAccountSearchCards=async(data)=>{
    var cardsArray = [
        {
            card: accountSearchModule.searchCard(data),
        }
    ];
    var isTpidEmpty = (data.tpId==undefined || data.tpId =='');
    var isSolutionAreaEmpty = (data.solutionArea==undefined || data.solutionArea =='');
    var isSalesPlayEmpty = (data.salesPlay==undefined || data.salesPlay =='');

    if(!(isTpidEmpty && isSalesPlayEmpty && isSolutionAreaEmpty) && ((isSolutionAreaEmpty && isSalesPlayEmpty ) || (isSalesPlayEmpty && isTpidEmpty) || (isTpidEmpty && isSolutionAreaEmpty))){
        var accountsSummaryCards = await accountSearchModule.fetchSearchAccountSummaryList(data);
        accountsSummaryCards.forEach(element=>{
            cardsArray.push({card:accountSummaryModule.getAccountSummaryCard(element,data.userId)});
        });
        if(accountsSummaryCards.length==0){
            cardsArray.push({card:invalidFilterAppliedCard()});
        }
    }else{
        cardsArray.push({card:invalidFilterAppliedCard()});
    }
    const res = {
        tab: {
            type: "continue",
            value: {
                cards: cardsArray
            },
        }
    };
    return res; 
}

module.exports = {
    fetchInitialAccountSeacrchCards,
    fetchAccountSearchCards
};