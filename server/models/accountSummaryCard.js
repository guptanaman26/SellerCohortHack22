var fs = require("fs");
const getAccountSummaryCard = (accountDetail, userId) => {
    const res2 = {
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
                                "width": 80,
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": accountDetail.accountName,
                                        "wrap": true,
                                        "size": "large",
                                        "weight": "Bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 20,
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": accountDetail.favourite?"https://www.pngpix.com/download/13890":"http://www.clker.com/cliparts/D/5/P/w/5/a/white-star-hi.png",
                                        "altText": "Cat",
                                        "selectAction": {
                                            "type": "Action.Submit",
                                            "data":{
                                                "loadView":"toggleFavourite",
                                                "tpId":accountDetail.tpId,
                                                "userId":userId
                                            }
                                        },
                                        "height": "25px",
                                        "horizontalAlignment": "Right"
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
                    "type": "Container",
                    "items": [
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "stretch",
                                    "style": "emphasis",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Total Contacts",
                                            "wrap": true
                                        },
                                        {
                                            "type": "TextBlock",
                                            "text": accountDetail.summaryData.totalContacts,
                                            "wrap": true
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "stretch",
                                    "style": "emphasis",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": "Total Interactions",
                                            "wrap": true
                                        },
                                        {
                                            "type": "TextBlock",
                                            "text": accountDetail.summaryData.totalInteractions,
                                            "wrap": true
                                        }
                                    ]
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
                            "width": "stretch",
                            "style": "emphasis",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": "Key Interactions",
                                    "wrap": true
                                },
                                {
                                    "type": "TextBlock",
                                    "text": accountDetail.summaryData.interactionsByDM,
                                    "wrap": true
                                }
                            ]
                        },
                        {
                            "type": "Column",
                            "width": "stretch",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": "Interaction Gain",
                                    "wrap": true
                                },
                                {
                                    "type": "TextBlock",
                                    "text": ((accountDetail.summaryData.interactionGain >= 0)? '▲' : '▼') + accountDetail.summaryData.interactionGain + "%",
                                    "wrap": true
                                }
                            ],
                            "style":  accountDetail.summaryData.interactionGain >= 0? 'good': 'attention'
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
                                "width": 70
                            },
                            {
                                "type": "Column",
                                "width": 30,
                                "items": [
                                    {
                                        "type": "Container",
                                        "items": [
                                            {
                                                "type": "ActionSet",
                                                "actions": [
                                                    {
                                                        "type": "Action.Submit",
                                                        "title": "More Details",
                                                        "data": {
                                                            "loadView": "accountDetails",
                                                            "tpId": accountDetail.tpId,
                                                            "solutionArea": "",
                                                            "salesPlay": "",
                                                            "accountName": accountDetail.accountName,
                                                            "userId":userId
                                                        }
                                                    }
                                                ]
                                            }
                                        ],
                                        "horizontalAlignment": "Right"
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }
        ]
    };
    return res2;
}

module.exports = {
    getAccountSummaryCard
};