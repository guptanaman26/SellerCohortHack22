const { StatusCodes } = require('botbuilder');
var fs = require("fs");
var LocalStorage = require('node-localstorage').LocalStorage;
let localStorage = new LocalStorage('./scratch');
    
const isFavourite=(tpId,userId)=>{
    let favTpidArray= localStorage.getItem(userId);
    if(favTpidArray==undefined){
        localStorage.setItem(userId,[]);
    }else{
        favTpidArray=favTpidArray.split(',');
        return favTpidArray.includes(tpId);
    }
    return false;
}

toggleFavourite=(data)=>{
    if(isFavourite(data.tpId,data.userId)){
        let favTpidArray= localStorage.getItem(data.userId).split(',');
        favTpidArray.splice(favTpidArray.indexOf(data.tpId),1);
        localStorage.setItem(data.userId,favTpidArray);
    }else{
        let favTpidArray= localStorage.getItem(data.userId).split(',');
        favTpidArray.push(data.tpId);
        localStorage.setItem(data.userId,favTpidArray);
    }
}

module.exports = {
  isFavourite,
  toggleFavourite
};