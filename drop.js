var Web3 = require('web3');
var Excel = require('exceljs');
var Tx = require('ethereumjs-tx');

//配置信息---------------------
var userAddr = '0xb5928F2BaF040Fa7665D0AF79e5517Cc6B7af398';
var nodeURI = 'https://ropsten.infura.io/zbI5uVZrIdl9VdoWDqMG';
// var nodeURI = 'https://mainnet.infura.io/zbI5uVZrIdl9VdoWDqMG';
var excelPath = '/Users/owen/Desktop/test-excel.xlsx';
var gasPrice = 20;
var gasLimit = 21000;
var pk = process.env.PK;
//-----------------------------

var privateKey = new Buffer.from(pk, 'hex');
var web3 = new Web3(Web3.providers.HttpProvider(nodeURI));
// var abi = [];
// var DosTokenContract = web3.eth.contract(abi);
// var dosAddr = "";
// var dosToken = DosTokenContract.at(dosAddr);

var curNum = 0;
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(excelPath)
    .then(async function() {
        var worksheet = workbook.getWorksheet(1);
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
                var callData = dosToken.transfer.getData(curAddr, curAmount);
                var nonce = await web3.eth.getTransactionCount(userAddr);
                var rawTx = {
                    nonce: nonce,
                    gasPrice: web3.utils.toHex(gasPrice),
                    gasLimit: web3.utils.toHex(gasLimit),
                    to: dosAddr,
                    value: '0x0',
                    data: callData
                }
                var tx = new Tx(rawTx);
                tx.sign(privateKey);
                var serializedTx = tx.serialize();
                web3.eth.sendRawTransaction("0x" + serializedTx.toString('hex'), function(err, hash) {
                    if (!err) {
                        curNum++;
                        console.log("行：" + rowNumber + ",地址：" + curAddr + ",金额：" + curAmount + ",哈希：" + hash);
                    } else
                        console.log("出错，行：" + rowNumber + ",地址：" + curAddr + ",金额：" + curAmount);
                    if (worksheet.rowCount == rowNumber) {
                        console.log("空投结束！！总人数：" + curNum);
                    }
                });
            }
        });
    });