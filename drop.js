var Web3 = require('web3');
var Excel = require('exceljs');
var Tx = require('ethereumjs-tx');

//config param---------------------
var pk = process.env.PK;
var userAddr = '';
// var nodeURI = 'https://ropsten.infura.io/zbI5uVZrIdl9VdoWDqMG';
var nodeURI = 'https://mainnet.infura.io/zbI5uVZrIdl9VdoWDqMG';
var excelPath = '/Users/owen/Desktop/test-excel.xlsx';
var gasPrice = 30;
var gasLimit = 60000;
var decimals = 1e18;
var abi = [{"constant":true,"inputs":[],"name":"name","outputs":[{"name":"","type":"string"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":false,"inputs":[],"name":"stop","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":false,"inputs":[{"name":"guy","type":"address"},{"name":"wad","type":"uint256"}],"name":"approve","outputs":[{"name":"","type":"bool"}],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":false,"inputs":[{"name":"owner_","type":"address"}],"name":"setOwner","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":true,"inputs":[],"name":"totalSupply","outputs":[{"name":"","type":"uint256"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":false,"inputs":[{"name":"src","type":"address"},{"name":"dst","type":"address"},{"name":"wad","type":"uint256"}],"name":"transferFrom","outputs":[{"name":"","type":"bool"}],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":true,"inputs":[],"name":"decimals","outputs":[{"name":"","type":"uint256"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":false,"inputs":[{"name":"guy","type":"address"},{"name":"wad","type":"uint256"}],"name":"mint","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":false,"inputs":[{"name":"wad","type":"uint256"}],"name":"burn","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":true,"inputs":[],"name":"manager","outputs":[{"name":"","type":"address"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":false,"inputs":[{"name":"_token","type":"address"},{"name":"_dst","type":"address"}],"name":"claimTokens","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":true,"inputs":[{"name":"src","type":"address"}],"name":"balanceOf","outputs":[{"name":"","type":"uint256"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":true,"inputs":[],"name":"stopped","outputs":[{"name":"","type":"bool"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":false,"inputs":[{"name":"authority_","type":"address"}],"name":"setAuthority","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":true,"inputs":[],"name":"owner","outputs":[{"name":"","type":"address"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":true,"inputs":[],"name":"symbol","outputs":[{"name":"","type":"string"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":false,"inputs":[{"name":"guy","type":"address"},{"name":"wad","type":"uint256"}],"name":"burn","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":false,"inputs":[{"name":"_newManager","type":"address"}],"name":"changeManager","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":false,"inputs":[{"name":"dst","type":"address"},{"name":"wad","type":"uint256"}],"name":"transfer","outputs":[{"name":"","type":"bool"}],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":false,"inputs":[],"name":"start","outputs":[],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":true,"inputs":[],"name":"authority","outputs":[{"name":"","type":"address"}],"payable":false,"stateMutability":"view","type":"function"},{"constant":false,"inputs":[{"name":"guy","type":"address"}],"name":"approve","outputs":[{"name":"","type":"bool"}],"payable":false,"stateMutability":"nonpayable","type":"function"},{"constant":true,"inputs":[{"name":"src","type":"address"},{"name":"guy","type":"address"}],"name":"allowance","outputs":[{"name":"","type":"uint256"}],"payable":false,"stateMutability":"view","type":"function"},{"inputs":[],"payable":false,"stateMutability":"nonpayable","type":"constructor"},{"payable":true,"stateMutability":"payable","type":"fallback"},{"anonymous":false,"inputs":[{"indexed":true,"name":"authority","type":"address"}],"name":"LogSetAuthority","type":"event"},{"anonymous":false,"inputs":[{"indexed":true,"name":"owner","type":"address"}],"name":"LogSetOwner","type":"event"},{"anonymous":true,"inputs":[{"indexed":true,"name":"sig","type":"bytes4"},{"indexed":true,"name":"guy","type":"address"},{"indexed":true,"name":"foo","type":"bytes32"},{"indexed":true,"name":"bar","type":"bytes32"},{"indexed":false,"name":"wad","type":"uint256"},{"indexed":false,"name":"fax","type":"bytes"}],"name":"LogNote","type":"event"},{"anonymous":false,"inputs":[{"indexed":true,"name":"from","type":"address"},{"indexed":true,"name":"to","type":"address"},{"indexed":false,"name":"value","type":"uint256"}],"name":"Transfer","type":"event"},{"anonymous":false,"inputs":[{"indexed":true,"name":"owner","type":"address"},{"indexed":true,"name":"spender","type":"address"},{"indexed":false,"name":"value","type":"uint256"}],"name":"Approval","type":"event"}];
var dosAddr = "0x70861e862E1Ac0C96f853C8231826e469eAd37B1";
//-----------------------------

var privateKey = new Buffer.from(pk, 'hex');
var web3 = new Web3(); 
web3.setProvider(new Web3.providers.HttpProvider(nodeURI));
var dosTemp = web3.eth.contract(abi);
var dosToken = dosTemp.at(dosAddr);

var workbook = new Excel.Workbook();
workbook.xlsx.readFile(excelPath)
    .then(async function() {
        var worksheet = workbook.getWorksheet(1);
        var nonce = await web3.eth.getTransactionCount(userAddr);
        worksheet.eachRow(function(row, rowNumber) {
            if (rowNumber != 1) {
                var curAddr;
                var curAmount;
                row.eachCell(function(cell, colNumber) {
                    if (colNumber == 1) {
                        curAddr = cell.value;
                    }
                    if (colNumber == 2) {
                        curAmount = cell.value;
                    }
                });
                var callData = dosToken.transfer.getData(curAddr, curAmount * decimals);               
                var rawTx = {
                    nonce: nonce++,
                    gasPrice: web3.toHex(gasPrice),
                    gasLimit: web3.toHex(gasLimit),
                    to: dosAddr,
                    value: '0x0',
                    data: callData
                }
                var tx = new Tx(rawTx);
                tx.sign(privateKey);
                var serializedTx = tx.serialize();
                web3.eth.sendRawTransaction("0x" + serializedTx.toString('hex'), function(err, hash) {
                    if (!err) {
                        console.log("行：" + rowNumber + "\t地址：" + curAddr + "\t金额：" + curAmount + "\t哈希：" + hash);
                    } else{
                        console.log("出错行：" + rowNumber + "\t地址：" + curAddr + "\t金额：" + curAmount + "\t错误信息：" + err);
                    }
                });
            }
        });
    });