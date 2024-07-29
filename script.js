function convertExcelDate(excelSerialDate) {
  const excelEpoch = new Date(Date.UTC(1900, 0, 1));
  const daysOffset = excelSerialDate - 2; // Excel dates start on 1/1/1900 with a leap year bug (Excel treats 1900 as a leap year)
  const convertedDate = new Date(
    excelEpoch.setUTCDate(excelEpoch.getUTCDate() + daysOffset)
  );
  return convertedDate.toISOString().split("T")[0]; // Format the date as YYYY-MM-DD
}

function loadData() {
  fetch("data.xlsx")
    .then((response) => response.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      const toDay = worksheet["U3"] ? worksheet["U3"].v : 0;

      document.getElementById("lastUpdated").textContent = `${convertExcelDate(
        toDay
      )}`;

      displayData(jsonData);
      calculateTotals(jsonData);
      console.log("Data refreshed");
    })
    .catch((error) =>
      console.error("Error fetching or parsing the .xlsx file:", error)
    );
}

function displayData(data) {
  const container = document.querySelector(".card-container");
  container.innerHTML = ""; // Clear previous data

  data.forEach((etf) => {
    const invested = parseFloat(etf["Money Invested"]) || 0;
    const shares = parseFloat(etf.Shares) || 0;
    const currentPrice = parseFloat(etf["Current Price"]) || 0;
    const currentValue = shares * currentPrice;
    const profit = currentValue - invested;
    const profitPercent = invested === 0 ? 0 : (profit / invested) * 100;

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
            <h2>${etf.Name}</h2>
            <p>Shares: ${shares}</p>
            <p>Money Invested: €${invested.toFixed(2)}</p>
            <p>Average Buy Price: €${parseFloat(
              etf["Average Buy Price"]
            ).toFixed(2)}</p>
            <p>Current Price: €${currentPrice.toFixed(2)}</p>
            <p>Total Current Value: €${currentValue.toFixed(2)}</p>
            <p>Total Profit: €${profit.toFixed(2)} (${profitPercent.toFixed(
      2
    )}%)</p>`;
    container.appendChild(card);
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

// Add event listener to the refresh button
document.getElementById("refreshData").addEventListener("click", loadData);
