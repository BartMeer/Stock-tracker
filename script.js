async function loadData() {
  try {
    const response = await fetch('data.xlsx');
    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    const today = sheet['U3']?.v || 0;

    displayDateTime(today);
    displayData(jsonData);
    calculateTotals(jsonData);
    console.log('Data refreshed');
  } catch (error) {
    console.error('Error fetching or parsing the .xlsx file:', error);
  }
}

function displayData(data) {
  const container = document.querySelector('.card-container');
  container.innerHTML = ''; // Clear previous data

  data.forEach((etf) => {
    const {
      Name,
      Shares,
      'Money Invested': invested,
      'Current Price': currentPrice,
      'Average Buy Price': avgBuyPrice,
    } = etf;

    const shares = parseFloat(Shares) || 0;
    const investedAmount = parseFloat(invested) || 0;
    const price = parseFloat(currentPrice) || 0;
    const avgPrice = parseFloat(avgBuyPrice) || 0;

    const currentValue = shares * price;
    const profit = currentValue - investedAmount;
    const profitPercent = investedAmount === 0 ? 0 : (profit / investedAmount) * 100;

    const card = document.createElement('div');
    card.className = 'card';
    card.innerHTML = `
      <h2>${Name}</h2>
      <p>Shares: ${numberFormatter.format(shares)}</p>
      <p>Money Invested: ${currencyFormatter.format(investedAmount)}</p>
      <p>Average Buy Price: ${currencyFormatter.format(avgPrice)}</p>
      <p>Current Price: ${currencyFormatter.format(price)}</p>
      <p>Total Current Value: ${currencyFormatter.format(currentValue)}</p>
      <p>Total Profit: ${currencyFormatter.format(profit)} (${numberFormatter.format(
      profitPercent.toFixed(2)
    )}%)</p>`;

    container.appendChild(card);
  });
}

function calculateTotals(data) {
  const totals = data.reduce(
    (acc, etf) => {
      const shares = parseFloat(etf.Shares) || 0;
      const invested = parseFloat(etf['Money Invested']) || 0;
      const currentPrice = parseFloat(etf['Current Price']) || 0;

      const currentValue = shares * currentPrice;
      const profit = currentValue - invested;

      acc.totalInvested += invested;
      acc.totalCurrentValue += currentValue;
      acc.totalProfit += profit;

      return acc;
    },
    { totalInvested: 0, totalCurrentValue: 0, totalProfit: 0 }
  );

  updateElement('totalInvested', 'Total Invested Money', totals.totalInvested);
  updateElement('totalCurrentValue', 'Total Current Value', totals.totalCurrentValue);
  updateElement('totalProfit', 'Total Profit', totals.totalProfit, totals.totalInvested, true);
}

// Automatically load data when the page loads
window.onload = loadData;
