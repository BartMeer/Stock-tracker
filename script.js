function loadData() {
  fetch("data.xlsx")
    .then((response) => response.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      displayData(jsonData);
      calculateTotals(jsonData);
    })
    .catch((error) =>
      console.error("Error fetching or parsing the .xlsx file:", error)
    );
}

function displayData(data) {
  const tbody = document.querySelector(".card-container");
  tbody.innerHTML = ""; // Clear previous data

  data.forEach((etf) => {
    const invested = parseFloat(etf["Money Invested"]) || 0;
    const shares = parseFloat(etf.Shares) || 0;
    const currentPrice = parseFloat(etf["Current Price"]) || 0;
    const currentValue = shares * currentPrice;
    const profit = currentValue - invested;
    const profitPercent = invested === 0 ? 0 : (profit / invested) * 100;

    const row = document.createElement("div");
    row.className = "card";
    row.innerHTML = `
            <h2>${etf.Name}</h2>
            <p>Shares: ${shares}</p>
            <p>Money Invested: €${invested.toFixed(2)}</p>
            <p>Average Buy Price: €${etf["Average Buy Price"].toFixed(2)}</p>
            <p>Current Price: €${currentPrice.toFixed(2)}</p>
            <p>Total Current Value: €${currentValue.toFixed(2)}</p>
            <p>Total Profit: €${profit.toFixed(2)} (${profitPercent.toFixed(
      2
    )}%)</p>`;
    tbody.appendChild(row);
  });
}

function calculateTotals(data) {
  let totalInvested = 0;
  let totalCurrentValue = 0;
  let totalProfit = 0;

  data.forEach((etf) => {
    const invested = parseFloat(etf["Money Invested"]) || 0;
    const shares = parseFloat(etf.Shares) || 0;
    const currentPrice = parseFloat(etf["Current Price"]) || 0;
    const currentValue = shares * currentPrice;
    const profit = currentValue - invested;

    totalInvested += invested;
    totalCurrentValue += currentValue;
    totalProfit += profit;
  });

  const profitPercent =
    totalInvested === 0 ? 0 : (totalProfit / totalInvested) * 100;

  document.getElementById(
    "totalInvested"
  ).textContent = `Total Invested Money: €${totalInvested.toFixed(2)}`;
  document.getElementById(
    "totalCurrentValue"
  ).textContent = `Total Current Value: €${totalCurrentValue.toFixed(2)}`;
  document.getElementById(
    "totalProfit"
  ).textContent = `Total Profit: €${totalProfit.toFixed(
    2
  )} (${profitPercent.toFixed(2)}%)`;
}

// Automatically load data when the page loads
window.onload = loadData;
