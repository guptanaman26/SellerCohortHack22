const { StatusCodes } = require('botbuilder');
var fs = require("fs");
var axios = require("axios");

const fetchAllSolutionAreaData= ()=>{
    const data =[
        {"name":"Business Applications"},
        {"name":"Infrastructure"},            
        {"name":"Digital and Application Innovation"},            
        {"name": "Data and AI"},            
           {"name": "Modern Work"},            
           {"name": "Security"},            
           {"name": "Microsoft Cloud and Industry"},
           {"name": "None"}
    ];
    const res = [];
    data.forEach(element => {
        res.push({
            "title": element.name,
            "value": element.name
        });
    });
    return res;
}
const fetchAllSalesPlayData=(solutionArea)=>{
    if(solutionArea==undefined || solutionArea == ''){
        return [{"title": "All","value": "All"}]
    }
    const data =[
        {
            "name": "Business Applications",
            "data": [
                {"name": "Accelerate Innovation with Low Code"},
                {"name": "Connected Sales and Marketing"},
                {"name": "Enable a Resilient and Sustainable Supply Chain"},
                {"name": "Modernize the Service Experience"},
                {"name": "Optimize Financial and Operating Models"},
                {"name": "Other"},
                {"name": "None"}
            ]
        },
        {
            "name": "Data and AI",
            "data": [
                {"name": "Enable Customer Success"},
                {"name": "Enable Developer Productivity and Accelerate Delivery"},
                {"name": "Enable Unified Data Governance"},
                {"name": "Innovate Across Hybrid and Edge with Azure Arc and IoT"},
                {"name": "Innovate and Scale with Cloud Native Apps"},
                {"name": "Innovate with AI and Cloud Scale Databases in Every App"},
                {"name": "Microsoft Azure Consumption Commitment (MACC)"},
                {"name": "Migrate and Modernize Your Data Estate"},
                {"name": "Migrate and Modernize Your Infrastructure and Workloads"},
                {"name": "Power Business Decisions with Cloud Scale Analytics"},
                {"name": "Other"},
                {"name": "None"}
            ]
        },
        {
            "name": "Infrastructure",
            "data": [
                {"name": "Enable Customer Success"},
                {"name": "Innovate Across Hybrid and Edge with Azure Arc and IoT"},
                {"name": "Innovate and Scale with Cloud Native Apps"},
                {"name": "Microsoft Azure Consumption Commitment (MACC)"},
                {"name": "Migrate and Modernize Your Infrastructure and Workloads"},
                {"name": "Modernize SAP on the Microsoft Cloud"},
                {"name": "Modernize Your Workloads with Azure at Any Scale with HPC Plus AI"},
                {"name": "Other"},
                {"name": "Protect Your Data and Ensure Business Resiliency with BCDR"}
            ]
        },
        {
            "name": "Modern Work",
            "data": [
                {"name": "Collaborative Apps"},
                {"name": "Converged Communications Teams Phone"},
                {"name": "Converged Communications Teams Rooms"},
                {"name": "Digital Workforce"},
                {"name": "Employee Experience"},
                {"name": "Frontline Workers"},
                {"name": "Migrate and Modernize Your Infrastructure and Workloads"},
                {"name": "NextGen Windows Experiences"},
                {"name": "Other"},
                {"name": "Refresh your devices"},
                {"name": "None"}
            ]
        },
        {
            "name": "Digital and Application Innovation",
            "data": [
                {"name": "Accelerate Innovation with Low Code"},
                {"name": "Build Your Games in the Cloud with Azure"},
                {"name": "Enable Customer Success"},
                {"name": "Enable Developer Productivity and Accelerate Delivery"},
                {"name": "Innovate Across Hybrid and Edge with Azure Arc and IoT"},
                {"name": "Innovate and Scale with Cloud Native Apps"},
                {"name": "Innovate with AI and Cloud Scale Databases in Every App"},
                {"name": "Microsoft Azure Consumption Commitment (MACC)"},
                {"name": "Modernize Enterprise Applications"},
                {"name": "Other"},
                {"name": "None"},
            ]
        },
        {
            "name": "Financial Services",
            "data": [
                {"name": "Combat Financial Crime"},
                {"name": "Deliver Differentiated Customer Experiences"},
                {"name": "Empower Employees Through Teamwork"},
                {"name": "Manage Risk Across the Organization"},
                {"name": "None"}
            ]
        },
        {
            "name": "Healthcare",
            "data": [
                {"name": "Empower Health Team Collaboration"},
                {"name": "Enhance Clinician Experiences"},
                {"name": "Enhance Patient Engagement"},
                {"name": "Improve Clinical and Operational Insights"}
            ]
        },
        {
            "name": "Retail",
            "data": [
                {"name": "Build a Real Time and Sustainable Supply Chain"},
                {"name": "Elevate the Shopping Experience"},
                {"name": "Empower Your Store Associate"},
                {"name": "Maximize the Value of Your Data"}
            ]
        },
        {
            "name": "Healthcare",
            "data": [
                {"name": "Empower Health Team Collaboration"},
                {"name": "Enhance Clinician Experiences"},
                {"name": "Enhance Patient Engagement"},
                {"name": "Improve Clinical and Operational Insights"}
            ]
        },
        {
            "name": "Security",
            "data": [
                {"name": "Defend Against Threats with SIEM Plus XDR"},
                {"name": "Enable Customer Success"},
                {"name": "Innovate Across Hybrid and Edge with Azure Arc and IoT"},
                {"name": "Mitigate Compliance and Privacy Risks"},
                {"name": "Other"},
                {"name": "Protect and govern sensitive data"},
                {"name": "Secure Identities and Access"},
                {"name": "Secure Multi Cloud Environments"},
                {"name": "None"}
            ]
        },
        {
            "name": "Sustainability",
            "data": [
                {"name": "Build a Sustainable IT Infrastructure"},
                {"name": "Create Sustainable Value Chains"},
                {"name": "Reduce Environmental Impact of Operations"},
                {"name": "Unify Data Intelligence"}
            ]
        },
        {
            "name": "Unified Support",
            "data": [
                {"name": "Build a strong Unified foundation"},
                {"name": "Ensure coverage for your key solutions"},
                {"name": "Modernize your digital estate with Unified Enterprise"}
            ]
        },
    ];
    const res = [];
    data.forEach(element => {
        if(element.name == solutionArea || solutionArea =='all'){
            element.data.forEach(element2=>{
                res.push({
                    "title": element2.name,
                    "value": element2.name
                });
            });
        }
    });
    return res;
}
const getFilterCard= (data)=>{
    var res={
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "ActionSet",
                        "actions": [
                            {
                                "type": "Action.Submit",
                                "title": "Home",
                                "data" : {
                                                "loadView" : "home"
                                            }
                            }
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": data.accountName,
                        "wrap": true,
                        "size": "Large",
                        "weight": "Bolder"
                    }
                ],
                "style": "accent"
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
                                        "id": "solutionArea",
                                        "choices": fetchAllSolutionAreaData(),
                                        "placeholder": "Solution Area",
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
                                        "id": "salesPlay",
                                        "choices": fetchAllSalesPlayData(data.solutionArea),
                                        "placeholder": "Sales Play",
                                        "value" : data.salesPlay
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
                                                "title": "Filter",
                                                "data":{
                                                        "loadView": "accountDetails",
                                                        "tpId": data.tpId,
                                                        "accountName": data.accountName,
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
                                                        "loadView": "clearAccountDetailsFilter",
                                                        "tpId": data.tpId,
                                                        "accountName": data.accountName,
                                                        "userId": data.userId
                                                    }
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": 60
                            }
                        ]
                    }
                ]
            }
        ]
    };
    return res;
}

const getColumnData=(data, colHeader, colKey)=>{
    var res = [{
        "type": "TextBlock",
        "weight": "Bolder",
        "text": colHeader
    }];
    data.forEach(element => {
        res.push({
            "type": "TextBlock",
            "separator": true,
            "text": element[colKey]
        });
    });
    return res;
}

const getDecisionMakerData = async (tpid, solutionArea, salesPlay) => {
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
      const response = await axios.get(
        `https://activitystore-ppe.trafficmanager.net/composite/api/v1/Activity/Details/CSP?Key=TPID&value=${tpid}&solutionArea=${solutionArea}&salesPlay=${salesPlay}`,
        config
      );
      let actualData = await response.data;
      let result = [];
      actualData["activities"].forEach(e => {
        if(e.activityDetail.marketingAudience && (e.activityDetail.marketingAudience.includes("BDM") || e.activityDetail.marketingAudience.includes("ITDM"))) return;
        let name = `${e.activityDetail.firstName} ${e.activityDetail.lastName} `;
        let activityDate = `${e.activityBase.activityDate}`;
        let phone = `${e.activityDetail.businessPhone}`;
        let email = `${e.activityDetail.email}`;
        let element = {
          "name": name,
          "activityDate": activityDate,
          "phone": phone,
          "email": email
        }
        result.push(element);
      });
      console.log(result);
      return result;
    } catch(err) {
      console.log(err.message)
        return (err.message);
    }
  }

const fetchDecisionMakerData=async(tpId, solutionArea, salesPlay)=>{
    var response=await getDecisionMakerData(tpId, solutionArea, salesPlay);
    return response;
}

const getDecisionMakerCard=(summaryData)=>{
    const res = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
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
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Name: ",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": summaryData.name,
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
                                        "text": "Lastest Interaction Date:",
                                        "wrap": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": summaryData.activityDate,
                                        "wrap": true
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
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Phone : " + summaryData.phone,
                                        "wrap": true
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
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Email : " + summaryData.email,
                                        "wrap": true
                                    }
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

function contains(arr, key1, val1, key2, val2) {
    for (var i = 0; i < arr.length; i++) {
      if (arr[i][key1] === val1 && arr[i][key2] === val2) return i;
    }
    return -1;
  } 
  const getAccountsData = async (tpid, solutionArea = null, salesPlay = null) => {
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
      if(solutionArea) url = `${url}&solutionArea=${solutionArea}`;
      if(salesPlay) url = `${url}&salesPlay=${salesPlay}`;
      const response = await axios.get(
        url,
        config
      );
      let actualData = await response.data;
      let result = [];
      actualData["activities"].forEach(e => {
        let solArea = `${e.activityDetail.solutionArea}`;
        let salesPlay = `${e.activityDetail.salesPlay}`;
        if (salesPlay == "undefined") return; 
        let index = contains(result, "solutionArea", solArea, "salesPlay", salesPlay); 
        if( index == -1) {
            let element = {
              "solutionArea": solArea,
              "salesPlay": salesPlay,
              "totalInteractions": 1,
              "totalInteractionsByDM": 0
            }
            result.push(element);
            if(e.activityDetail.marketingAudience &&
                (e.activityDetail.marketingAudience.includes("BDM") 
                || e.activityDetail.marketingAudience.includes("ITDM"))
              ) 
              result[result.length-1]["totalInteractionsByDM"]++;
          } else {
            result[index]["totalInteractions"]++;
            if(e.activityDetail.marketingAudience &&
                (e.activityDetail.marketingAudience.includes("BDM") 
                || e.activityDetail.marketingAudience.includes("ITDM"))
              ) 
              result[index]["totalInteractionsByDM"]++;
          }
      });
      console.log("s")
      console.log(result);
      return result;
    } catch(err) {
      console.log(err.message)
        return (err.message);
    }
  }

const fetchAccountDetails=async(tpId, solutionArea, salesPlay)=>{
    var accountDetails = await   getAccountsData(tpId, solutionArea, salesPlay);
    return accountDetails;
}


const getInteractedDetails = async(summaryData)=>{
    const data = await fetchAccountDetails(summaryData.tpId, summaryData?.solutionArea, summaryData?.salesPlay);
    var activeColumns = ['totalInteractions', 'totalInteractionsByDM'];
    if(summaryData.solutionArea==undefined || summaryData.solutionArea==''){
        activeColumns.push('solutionArea');
    }
    if(summaryData.salesPlay==undefined || summaryData.salesPlay ==''){
        activeColumns.push('salesPlay');
    }
    const res ={
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Interaction Details",
                        "wrap": true,
                        "weight": "Bolder",
                        "size": "Large"
                    }
                ],
                "style": "emphasis"
            },
            {
                "type": "ColumnSet",
                "columns": fetchInteractionColumns(data,activeColumns)
            }
        ]
    };
      
    return res;
}

const fetchInteractionColumns=(data,activeColumns)=>{
    var interactionColumns=[];
    if(activeColumns.includes('solutionArea')){
        interactionColumns.push({
            "type": "Column",
            "width": "stretch",
            "items": getColumnData(data,"Solution Area","solutionArea")
        });
    }
    if(activeColumns.includes('salesPlay')){
        interactionColumns.push({
            "type": "Column",
            "width": "stretch",
            "items": getColumnData(data,"Sales Play","salesPlay")
        });
    }
    interactionColumns.push({
        "type": "Column",
        "width": "stretch",
        "items": getColumnData(data,"Total Interactions","totalInteractions")
    });
    interactionColumns.push({
        "type": "Column",
        "width": "stretch",
        "items": getColumnData(data,"Key Interactions","totalInteractionsByDM")
    });
    return interactionColumns;
}

module.exports = {
    getFilterCard,
    getInteractedDetails,
    getDecisionMakerCard,
    fetchAllSolutionAreaData,
    fetchAllSalesPlayData,
    fetchDecisionMakerData
};