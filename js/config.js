const CONFIG = {
  API_KEY: 'AIzaSyAB6yjxTB0TNbEk2C68aOP5u0IkdmK12tg',
  SHEET_ID: '1FF4odviKZ2LnRvPf8ltM0o_jxM0ZHuJHBlkQCjC3sxA',
  SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbziDXIkJa_VIXVJpRnwv5aYDq425OU5O1vkDvMXEDmzj5KAzg80PJQFtN5DKOmlv0qp/exec'
};

const USERS = {
  m: { password: 'm', role: 'Manager', nickname: 'Manager' },
  t: { password: 't', role: 'Teller', nickname: 'Teller' },
  a: { password: 'a', role: 'Accountant', nickname: 'Accountant' },
  u1: { password: 'u1', role: 'User', nickname: 'Sale 1' },
  u2: { password: 'u2', role: 'User', nickname: 'Sale 2' }
};

const FIXED_PRODUCTS = [
  { id: 'G01', name: '10 บาท', unit: 'แท่ง', weight: 10 },
  { id: 'G02', name: '5 บาท', unit: 'แท่ง', weight: 5 },
  { id: 'G03', name: '2 บาท', unit: 'แท่ง', weight: 2 },
  { id: 'G04', name: '1 บาท', unit: 'แท่ง', weight: 1 },
  { id: 'G05', name: '2 สลึง', unit: 'แท่ง', weight: 0.25 },
  { id: 'G06', name: '1 สลึง', unit: 'แท่ง', weight: 0.125 },
  { id: 'G07', name: '1 กรัม', unit: 'แท่ง', weight: 1/15 }
];

const GOLD_WEIGHTS = {
  'G01': 10, 
  'G02': 5, 
  'G03': 2, 
  'G04': 1,
  'G05': 0.25, 
  'G06': 0.125,
  'G07': 1/15
};

const PREMIUM_PRODUCTS = ['G05', 'G06', 'G07'];
let PREMIUM_PER_PIECE = 100000;

const EXCHANGE_FEES = {
  'G01': 1000000,
  'G02': 500000,
  'G03': 200000,
  'G04': 100000,
  'G05': 25000,
  'G06': 12500,
  'G07': 10000
};

let currentUser = null;
let currentPricing = {
  sell1Baht: 0,
  buyback1Baht: 0
};

let currentPriceRates = {
  thbSell: 0,
  usdSell: 0,
  thbBuy: 0,
  usdBuy: 0
};

let sellSortOrder = 'desc';
let tradeinOldCounter = 0;
let tradeinNewCounter = 0;
let exchangeOldCounter = 0;
let exchangeNewCounter = 0;

let currentTradeinData = null;
let currentExchangeData = null;
let currentReconcileType = null;
let currentReconcileData = {};
let currentPaymentData = null;
