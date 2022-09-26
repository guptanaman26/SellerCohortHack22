const { StatusCodes } = require('botbuilder');
var fs = require("fs");
var accountDetailsCard = require('./accountDetailsCard');
var tpidData = require ('./sellerAccountMapping');
var adaptiveModule = require('./adaptiveCard');

const getUserAssociatedAccounts= (userId) =>{
    var response=[];
    tpidData.TpidAccountMapping.forEach(element=>{
        if(userId==element.SellerID){
            response.push({
                "title":element.AccountName,
                "value":element.TPID
            });
        }
    });
    return response;
}

const searchCard= (data)=>{
    var res = {
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
                                "width": 20,
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "Home",
                                                "data":{
                                                    "loadView":"home"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 80
                            }
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Input.ChoiceSet",
                                        "choices": getUserAssociatedAccounts(data.userId),
                                        "placeholder": "Select Account Name",
                                        "id":"tpId",
                                        "value": data.tpId
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Input.ChoiceSet",
                                        "choices": accountDetailsCard.fetchAllSolutionAreaData(),
                                        "placeholder": "Select Solution Area",
                                        "id": "solutionArea",
                                        "value": data.solutionArea
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "Input.ChoiceSet",
                                        "choices": accountDetailsCard.fetchAllSalesPlayData('all'),
                                        "placeholder": "Select Sales Play",
                                        "id":"salesPlay",
                                        "value": data.salesPlay
                                    }
                                ]
                            }
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": 20,
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "Search",
                                                "data":{
                                                    "loadView" : "searchResults",
                                                    "userId": data.userId
                                                }
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 20,
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "Clear",
                                                "data": {
                                                        "loadView": "initialSearch",
                                                        "userId": data.userId
                                                    }
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 60,
                                "items":[
                                    
                                ]
                            }
                        ]
                    }
                ]
            }
        ]
    };
    return res;
}

//there are three filters
const fetchSearchAccountSummaryList=async(data)=>{
    let tpIdArray=[];
    if(data.tpId!=undefined && data.tpId!=''){
        tpIdArray.push(data.tpId);
    }else if(data.solutionArea!=undefined && data.solutionArea!=''){
        tpidData.TpidAccountMapping.forEach(element=>{
            if(element.SolutionArea==data.solutionArea && element.SellerID == data.userId){
                tpIdArray.push(element.TPID);
            }
        });
    } else if(data.salesPlay!=undefined && data.salesPlay!=''){
        tpidData.TpidAccountMapping.forEach(element=>{
            if(element.SalesPlay==data.salesPlay && element.SellerID == data.userId){
                tpIdArray.push(element.TPID);
            }
        });
    }
    
    return await adaptiveModule.getAccountSummaryList(tpIdArray,data.userId);
}

module.exports = {
    searchCard,
    fetchSearchAccountSummaryList
}