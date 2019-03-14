var Web3 = require('web3');
var Excel = require('exceljs');
var Tx = require('ethereumjs-tx');
var ethUtils = require('ethereumjs-util');
var assert = require('assert');
var abiJson = require('./DOSToken.json');

//config param---------------------
var pk = process.env.PK;
var excelPath = process.env.ExcelPath;
assert(pk !== undefined && pk.length == 64, "Please export private key in hex format without leading '0x'");
assert(excelPath !== undefined, "Please specify path to excel file containing airdrop information");
var pkBuf = Buffer.from(pk, 'hex');
var fromAddr = '0x' + ethUtils.privateToAddress(pkBuf).toString('hex');
// var nodeURI = 'https://ropsten.infura.io/zbI5uVZrIdl9VdoWDqMG';
var nodeURI = 'https://mainnet.infura.io/zbI5uVZrIdl9VdoWDqMG';
var gasPrice = 10 * 1e9;  // 10Gwei
var gasLimit = 70000;
var decimals = 1e18;
var dosAddr = "0x70861e862E1Ac0C96f853C8231826e469eAd37B1";
//-----------------------------

var web3 = new Web3(new Web3.providers.HttpProvider(nodeURI)); 
var dosToken = web3.eth.contract(abiJson).at(dosAddr);

var workbook = new Excel.Workbook();
workbook.xlsx.readFile(excelPath)
    .then(async function() {
        var worksheet = workbook.getWorksheet(1);
        var nonce = await web3.eth.getTransactionCount(fromAddr);
        worksheet.eachRow(function(row, rowNumber) {
            if (rowNumber != 1) {
                var droppedAddr;
                var droppedAmount;
                row.eachCell(function(cell, colNumber) {
                    if (colNumber == 1) {
                        droppedAddr = cell.value;
                    }
                    if (colNumber == 2) {
                        droppedAmount = cell.value;
                    }
                });
                var callData = dosToken.transfer.getData(droppedAddr, droppedAmount * decimals);               
                var rawTx = {
                    nonce: nonce++,
                    gasPrice: web3.toHex(gasPrice),
                    gasLimit: web3.toHex(gasLimit),
                    to: dosAddr,
                    value: '0x0',
                    data: callData
                };
                var tx = new Tx(rawTx);
                tx.sign(pkBuf);
                var serializedTx = tx.serialize();
                web3.eth.sendRawTransaction("0x" + serializedTx.toString('hex'), function(err, hash) {
                    if (!err) {
                        console.log("Line: " + rowNumber + "\tDropped Address: " + droppedAddr + "\tAmount: " + droppedAmount + "\tTxHash: " + hash);
                    } else{
                        console.log("Err line: " + rowNumber + "\tAddress: " + droppedAddr + "\tAmount: " + droppedAmount + "\tErr message: " + err);
                    }
                });
            }
        });
    });
