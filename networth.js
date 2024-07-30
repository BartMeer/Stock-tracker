function loadNetWorthData() {
  fetch('data.xlsx')
    .then((response) => response.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const toDay = getCellValue(sheet['U3']);

      displayDateTime(toDay);

      // Define cell ranges for each person
      const row2 = {
        netWorth: getCellValue(sheet['G2']),
        bolero: getCellValue(sheet['H2']),
        cash: getCellValue(sheet['I2']),
        pension: getCellValue(sheet['J2']),
      };

      const row3 = {
        netWorth: getCellValue(sheet['G3']),
        bolero: getCellValue(sheet['H3']),
        cash: getCellValue(sheet['I3']),
        kbc: getCellValue(sheet['J3']),
      };

      // Calculate totals
      const totalNetWorth = row2.netWorth + row3.netWorth;
      const totalBolero = row2.bolero + row3.bolero;
      const totalCash = row2.cash + row3.cash;
      const totalPension = row2.pension;
      const totalKBC = row3.kbc;

      // Update HTML with data for Bart
      updateElement('netWorth1', 'Net Worth', row2.netWorth, totalNetWorth, true);
      updateElement('allocationBolero1', 'Bolero', row2.bolero, row2.netWorth, true);
      updateElement('allocationCash1', 'Cash', row2.cash, row2.netWorth, true);
      updateElement('allocationPension1', 'Pension Savings', row2.pension, row2.netWorth, true);

      // Update HTML with data for Jolien
      updateElement('netWorth2', 'Net Worth', row3.netWorth, totalNetWorth, true);
      updateElement('allocationBolero2', 'Bolero', row3.bolero, row3.netWorth, true);
      updateElement('allocationCash2', 'Cash', row3.cash, row3.netWorth, true);
      updateElement('allocationKBC', 'KBC Stocks', row3.kbc, row3.netWorth, true);

      // Update HTML with combined data
      updateElement('combinedNetWorth', 'Net Worth', totalNetWorth);
      updateElement('combinedAllocationBolero', 'Bolero', totalBolero, totalNetWorth, true);
      updateElement('combinedAllocationCash', 'Cash', totalCash, totalNetWorth, true);
      updateElement(
        'combinedAllocationPension',
        'Pension Savings',
        totalPension,
        totalNetWorth,
        true
      );
      updateElement('combinedAllocationKBC', 'KBC Stocks', totalKBC, totalNetWorth, true);
    })
    .catch((error) => console.error('Error fetching or parsing the .xlsx file:', error));
}

window.onload = loadNetWorthData;
