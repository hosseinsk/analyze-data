// Process Excel Data - کامل و دقیق
const ExcelJS = require('exceljs');

// دریافت داده‌های استخراج شده از قدم قبل
const queryInfo = $input.first();
const plateNumbers = queryInfo.json.plateNumbers; // آرایه همه پلاک‌ها
const startDateStr = queryInfo.json.startDate; // "1404/04/01"
const endDateStr = queryInfo.json.endDate;     // "1404/04/30"

// تبدیل تاریخ شمسی به فرمت YYYYMMDD برای فیلتر
function convertPersianDate(dateStr) {
  const parts = dateStr.split('/');
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]).toString().padStart(2, '0');
  const day = parseInt(parts[2]).toString().padStart(2, '0');
  return parseInt(`${year}${month}${day}`);
}

// تبدیل YYYYMMDD به تاریخ قابل خواندن
function formatDateForDisplay(yyyymmdd) {
  const dateStr = yyyymmdd.toString();
  const year = dateStr.substring(0, 4);
  const month = dateStr.substring(4, 6);
  const day = dateStr.substring(6, 8);
  return `${year}/${month}/${day}`;
}

// تبدیل فرمت YYYY/MM/DD به YYYYMMDD
function convertSlashDateToNumber(dateStr) {
  if (!dateStr || dateStr === 'NaN' || typeof dateStr !== 'string') {
    return null;
  }
  const parts = dateStr.split('/');
  if (parts.length !== 3) return null;
  
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]).toString().padStart(2, '0');
  const day = parseInt(parts[2]).toString().padStart(2, '0');
  return parseInt(`${year}${month}${day}`);
}

const startDate = convertPersianDate(startDateStr);
const endDate = convertPersianDate(endDateStr);

// دریافت فایل Excel
const binaryData = $('Download Excel').first().binary;
const workbook = new ExcelJS.Workbook();

try {
  await workbook.xlsx.load(binaryData);
  
  // شیت‌های مورد نیاز
  const dailyFuelSheet = workbook.getWorksheet('سوخت روزانه');
  const fuelRecordSheet = workbook.getWorksheet('ثبت اطلاعات سوخت');
  const dataSheet = workbook.getWorksheet('data');
  
  if (!dailyFuelSheet || !fuelRecordSheet || !dataSheet) {
    return [{
      json: {
        error: true,
        message: 'یکی از شیت‌های مورد نیاز یافت نشد. شیت‌های لازم: سوخت روزانه، ثبت اطلاعات سوخت، data'
      }
    }];
  }

  // ===== خواندن شیت "سوخت روزانه" =====
  const dailyFuelData = [];
  dailyFuelSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // رد کردن هدر
      const dateValue = row.getCell(5).value; // ستون E
      const plateValue = row.getCell(7).value; // ستون G
      const fuelTankValue = row.getCell(8).value; // ستون H
      const kmValue = row.getCell(9).value; // ستون I
      
      if (dateValue && plateValue) {
        const dateNum = parseInt(dateValue.toString());
        const plateStr = plateValue.toString();
        
        dailyFuelData.push({
          date: dateNum,
          plate: plateStr,
          fuelTank: parseFloat(fuelTankValue || 0),
          kilometers: parseFloat(kmValue || 0)
        });
      }
    }
  });

  // ===== خواندن شیت "ثبت اطلاعات سوخت" =====
  const fuelRecordData = [];
  fuelRecordSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // رد کردن هدر
      const dateValue = row.getCell(5).value; // ستون E
      const plateValue = row.getCell(11).value; // ستون K
      const fuelAmountValue = row.getCell(19).value; // ستون S
      
      if (dateValue && plateValue && fuelAmountValue) {
        const dateNum = parseInt(dateValue.toString());
        const plateStr = plateValue.toString();
        
        fuelRecordData.push({
          date: dateNum,
          plate: plateStr,
          fuelAmount: parseFloat(fuelAmountValue)
        });
      }
    }
  });

  // ===== خواندن شیت "data" =====
  const serviceData = [];
  dataSheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // رد کردن هدر
      const dateValue = row.getCell(5).value; // ستون E
      const plateValue = row.getCell(19).value; // ستون S
      const originValue = row.getCell(8).value; // ستون H
      
      if (dateValue && plateValue) {
        const dateStr = dateValue.toString();
        const dateNum = convertSlashDateToNumber(dateStr);
        const plateStr = plateValue.toString();
        const originStr = (originValue || '').toString().trim();
        
        if (dateNum) {
          serviceData.push({
            date: dateNum,
            plate: plateStr,
            origin: originStr
          });
        }
      }
    }
  });

  // ===== تولید گزارش برای همه پلاک‌ها =====
  const reports = [];
  let rowNumber = 1;

  // استخراج همه پلاک‌های منحصر به فرد
  const allPlates = new Set();
  dailyFuelData.forEach(item => allPlates.add(item.plate));
  fuelRecordData.forEach(item => allPlates.add(item.plate));
  serviceData.forEach(item => allPlates.add(item.plate));

  for (const plateNumber of Array.from(allPlates).sort()) {
    // A: ردیف
    const row = rowNumber++;
    
    // B: پلاک کامیون
    const plate = plateNumber;
    
    // C: تاریخ شروع
    const startDateDisplay = startDateStr;
    
    // D: تاریخ پایان
    const endDateDisplay = endDateStr;
    
    // E: مجموع سوخت‌گیری در بازه
    const totalFuelReceived = fuelRecordData
      .filter(item => 
        item.plate === plateNumber && 
        item.date >= startDate && 
        item.date <= endDate
      )
      .reduce((sum, item) => sum + item.fuelAmount, 0);
    
    // فیلتر داده‌های سوخت روزانه برای این پلاک در بازه
    const plateDailyData = dailyFuelData
      .filter(item => 
        item.plate === plateNumber && 
        item.date >= startDate && 
        item.date <= endDate
      )
      .sort((a, b) => a.date - b.date);
    
    // F: کیلومتر کارکرد
    let kilometersWorked = 0;
    let startKm = 0;
    let endKm = 0;
    
    if (plateDailyData.length > 0) {
      // اولین و آخرین کیلومتر در بازه
      const validKmData = plateDailyData.filter(item => item.kilometers > 0);
      
      if (validKmData.length > 0) {
        startKm = validKmData[0].kilometers;
        endKm = validKmData[validKmData.length - 1].kilometers;
        kilometersWorked = Math.max(0, endKm - startKm);
      }
    }
    
    // G: سوخت مصرفی دوره
    let fuelConsumedPeriod = 0;
    let startFuelTank = 0;
    let endFuelTank = 0;
    
    if (plateDailyData.length > 0) {
      const validFuelData = plateDailyData.filter(item => item.fuelTank > 0);
      
      if (validFuelData.length > 0) {
        startFuelTank = validFuelData[0].fuelTank;
        endFuelTank = validFuelData[validFuelData.length - 1].fuelTank;
        fuelConsumedPeriod = (startFuelTank - endFuelTank) + totalFuelReceived;
      } else {
        fuelConsumedPeriod = totalFuelReceived;
      }
    }
    
    // فیلتر داده‌های سرویس برای این پلاک در بازه
    const plateServiceData = serviceData
      .filter(item => 
        item.plate === plateNumber && 
        item.date >= startDate && 
        item.date <= endDate
      );
    
    // H: تعداد سرویس شهری
    const cityServices = plateServiceData
      .filter(item => 
        item.origin === 'شهرک صنعتی' || 
        item.origin === 'شمال پالایشگاه'
      ).length;
    
    // I: تعداد سرویس شهرستان
    const provinceServices = plateServiceData.length - cityServices;
    
    // J: سوخت مصرفی شهری بر اساس فرمول
    const cityFuelConsumption = cityServices * 25;
    
    // K: سوخت مصرفی شهرستان بر اساس فرمول
    const abadanServices = plateServiceData
      .filter(item => 
        item.origin !== 'شهرک صنعتی' && 
        item.origin !== 'شمال پالایشگاه' &&
        item.origin === 'آبادان'
      ).length;
    
    const asaluyehServices = plateServiceData
      .filter(item => 
        item.origin !== 'شهرک صنعتی' && 
        item.origin !== 'شمال پالایشگاه' &&
        item.origin === 'عسلویه'
      ).length;
    
    const otherProvinceServices = provinceServices - abadanServices - asaluyehServices;
    
    const provinceFuelConsumption = 
      (abadanServices * 800) + 
      (asaluyehServices * 400) + 
      (otherProvinceServices * 600); // فرض کنیم بقیه 600 لیتر
    
    // L: مجموع سوخت مصرفی بر اساس سرویس
    const totalServiceBasedFuel = cityFuelConsumption + provinceFuelConsumption;
    
    // N: سوخت مصرفی بر اساس کیلومتر
    const kmBasedFuelConsumption = kilometersWorked * 0.6;
    
    // M: اختلاف مصرف سوخت
    const fuelConsumptionDifference = kmBasedFuelConsumption - totalServiceBasedFuel;
    
    // اضافه کردن به گزارش
    reports.push({
      A_ردیف: row,
      B_پلاک: plate,
      C_تاریخ_شروع: startDateDisplay,
      D_تاریخ_پایان: endDateDisplay,
      E_مجموع_سوخت_گیری: Math.round(totalFuelReceived * 10) / 10,
      F_کیلومتر_کارکرد: Math.round(kilometersWorked),
      G_سوخت_مصرفی_دوره: Math.round(fuelConsumedPeriod * 10) / 10,
      H_تعداد_سرویس_شهری: cityServices,
      I_تعداد_سرویس_شهرستان: provinceServices,
      J_سوخت_شهری_فرمول: Math.round(cityFuelConsumption * 10) / 10,
      K_سوخت_شهرستان_فرمول: Math.round(provinceFuelConsumption * 10) / 10,
      L_مجموع_سوخت_سرویس: Math.round(totalServiceBasedFuel * 10) / 10,
      M_اختلاف_مصرف: Math.round(fuelConsumptionDifference * 10) / 10,
      N_سوخت_کیلومتری: Math.round(kmBasedFuelConsumption * 10) / 10,
      
      // اطلاعات تکمیلی
      startKm: startKm,
      endKm: endKm,
      startFuelTank: startFuelTank,
      endFuelTank: endFuelTank,
      abadanServices: abadanServices,
      asaluyehServices: asaluyehServices,
      otherProvinceServices: otherProvinceServices,
      totalRecords: plateDailyData.length
    });
  }

  // مرتب‌سازی بر اساس پلاک
  reports.sort((a, b) => a.B_پلاک.localeCompare(b.B_پلاک));

  return [{
    json: {
      success: true,
      startDate: startDateStr,
      endDate: endDateStr,
      totalPlates: reports.length,
      reports: reports,
      summary: {
        totalFuelReceived: reports.reduce((sum, r) => sum + r.E_مجموع_سوخت_گیری, 0),
        totalKilometers: reports.reduce((sum, r) => sum + r.F_کیلومتر_کارکرد, 0),
        totalCityServices: reports.reduce((sum, r) => sum + r.H_تعداد_سرویس_شهری, 0),
        totalProvinceServices: reports.reduce((sum, r) => sum + r.I_تعداد_سرویس_شهرستان, 0)
      }
    }
  }];

} catch (error) {
  return [{
    json: {
      error: true,
      message: `خطا در پردازش فایل Excel: ${error.message}`,
      stack: error.stack
    }
  }];
}
