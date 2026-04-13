const CONFIG = {
  DIESEL_PRICE_PER_LITER: 73,
  CONSUMPTION_PER_100_KM: 11,
  COST_PER_KM: 8.1,
  LEASING_SHARE: 0.2,
  MAINTENANCE_SHARE: 0.4,
  DRIVER_SHARE: 0.4,
  VAT_DRIVER_TAX_SHARE: 0.19,
};

const SHEET_NAME = 'Trips';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Заявки')
    .addItem('Открыть панель заявок', 'showTripsSidebar')
    .addItem('Внести поездку', 'showTripForm')
    .addToUi();

  ensureTripsSheet_();
}

function showTripsSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Панель «Заявки»');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showTripForm() {
  const html = HtmlService.createTemplateFromFile('TripForm')
    .evaluate()
    .setWidth(900)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Внести поездку');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function fetchTrips() {
  const sheet = ensureTripsSheet_();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map((row) => ({
    id: row[0],
    createdAt: row[1],
    mainRoute: row[2],
    distanceKm: row[3],
    emptyReturn: row[4],
    totalRevenue: row[5],
    totalFuelCost: row[6],
    netAmount: row[7],
    displayTotal: row[5],
  })).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

function getTripDetails(tripId) {
  const sheet = ensureTripsSheet_();
  const rows = sheet.getDataRange().getValues();
  const match = rows.slice(1).find((row) => String(row[0]) === String(tripId));
  if (!match) {
    throw new Error('Поездка не найдена');
  }

  const cargos = JSON.parse(match[14] || '[]');
  return {
    id: match[0],
    createdAt: match[1],
    mainRoute: match[2],
    distanceKm: match[3],
    emptyReturn: match[4],
    totalRevenue: match[5],
    totalFuelCost: match[6],
    netAmount: match[7],
    leasingAmount: match[8],
    maintenanceAmount: match[9],
    driverAmount: match[10],
    driverTaxAmount: match[11],
    emptyReturnFuelAmount: match[12],
    companyAfterAdjustments: match[13],
    cargos,
  };
}

function saveTrip(payload) {
  validatePayload_(payload);

  const calc = calculateTrip_(payload);
  const sheet = ensureTripsSheet_();
  const tripId = Utilities.getUuid();

  sheet.appendRow([
    tripId,
    new Date(),
    payload.mainRoute,
    Number(payload.distanceKm),
    Boolean(payload.emptyReturn),
    calc.totalRevenue,
    calc.totalFuelCost,
    calc.netAmount,
    calc.leasingAmount,
    calc.maintenanceAmount,
    calc.driverAmount,
    calc.driverTaxAmount,
    calc.emptyReturnFuelAmount,
    calc.companyAfterAdjustments,
    JSON.stringify(payload.cargos),
  ]);

  return {
    success: true,
    tripId,
    ...calc,
  };
}

function calculateTrip_(payload) {
  const distanceKm = Number(payload.distanceKm);
  const cargos = payload.cargos;

  const totalRevenue = cargos.reduce((sum, cargo) => sum + Number(cargo.price), 0);
  const totalFuelCost = round2_(distanceKm * CONFIG.COST_PER_KM);
  const netAmount = round2_(totalRevenue - totalFuelCost);

  const leasingAmount = round2_(netAmount * CONFIG.LEASING_SHARE);
  const maintenanceAmount = round2_(netAmount * CONFIG.MAINTENANCE_SHARE);

  const baseDriverAmount = round2_(netAmount * CONFIG.DRIVER_SHARE);
  const driverTaxAmount = round2_(
    cargos
      .filter((cargo) => cargo.paymentType === 'with_vat')
      .reduce((sum, cargo) => {
        const cargoNet = Number(cargo.price) - totalFuelCost * (Number(cargo.price) / totalRevenue);
        return sum + cargoNet * CONFIG.DRIVER_SHARE * CONFIG.VAT_DRIVER_TAX_SHARE;
      }, 0)
  );

  let driverAmount = round2_(baseDriverAmount - driverTaxAmount);
  let companyAfterAdjustments = round2_(leasingAmount + maintenanceAmount);
  let emptyReturnFuelAmount = 0;

  if (payload.emptyReturn) {
    emptyReturnFuelAmount = totalFuelCost;
    const halfFuel = round2_(emptyReturnFuelAmount / 2);
    driverAmount = round2_(driverAmount - halfFuel);
    companyAfterAdjustments = round2_(companyAfterAdjustments - halfFuel);
  }

  return {
    totalRevenue: round2_(totalRevenue),
    totalFuelCost,
    netAmount,
    leasingAmount,
    maintenanceAmount,
    driverAmount,
    driverTaxAmount,
    emptyReturnFuelAmount,
    companyAfterAdjustments,
  };
}

function validatePayload_(payload) {
  if (!payload || !payload.mainRoute) {
    throw new Error('Укажите основной маршрут');
  }

  if (!payload.distanceKm || Number(payload.distanceKm) <= 0) {
    throw new Error('Укажите расстояние в км');
  }

  if (!Array.isArray(payload.cargos) || payload.cargos.length === 0) {
    throw new Error('Добавьте хотя бы один грузовой маршрут');
  }

  payload.cargos.forEach((cargo, index) => {
    if (!cargo.route) {
      throw new Error(`Заполните маршрут груза №${index + 1}`);
    }
    if (!cargo.price || Number(cargo.price) <= 0) {
      throw new Error(`Заполните стоимость груза №${index + 1}`);
    }
    if (!['cash', 'without_vat', 'with_vat'].includes(cargo.paymentType)) {
      throw new Error(`Выберите тип оплаты для груза №${index + 1}`);
    }
  });
}

function ensureTripsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'ID',
      'Дата',
      'Основной маршрут',
      'Км',
      'Холостой обратный путь',
      'Выручка',
      'Топливо',
      'Остаток после топлива',
      'Лизинг (20%)',
      'Ремонт/обслуживание (40%)',
      'Водитель',
      'Налог с НДС для водителя (19%)',
      'Топливо на холостой путь',
      'Итого компании после корректировок',
      'Грузы JSON',
    ]);
    sheet.getRange(1, 1, 1, 15).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function round2_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}
