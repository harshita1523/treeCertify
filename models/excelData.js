// Inside models/excelData.js
const mongoose = require('mongoose');

const dynamicSchema = new mongoose.Schema({}, { strict: false });

const ExcelData = mongoose.model('ExcelData', dynamicSchema);

module.exports = ExcelData;
