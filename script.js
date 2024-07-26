function loadData() {
  fetch("data.xlsx")
    .then((response) => response.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      displayData(jsonData);

      // Read net worth from cell G2
      const netWorthCell = worksheet["G2"];
      const netWorth = netWorthCell ? netWorthCell.v : 0;

      document.getElementById(
        "netWorth"
      ).textContent = `Net Worth: €${netWorth.toFixed(2)}`;
    })
    .catch((error) =>
      console.error("Error fetching or parsing the .xlsx file:", error)
    );
}

function displayData(data) {
  const container = document.getElementById("etfCards");
  container.innerHTML = ""; // Clear previous data

  let totalInvested = 0;
  let totalCurrentValue = 0;

  data.forEach((etf) => {
    const invested = parseFloat(etf["Money Invested"]) || 0;
    const shares = parseFloat(etf.Shares) || 0;
    const currentPrice = parseFloat(etf["Current Price"]) || 0;
    const averageBuyPrice = parseFloat(etf["Average Buy Price"]) || 0;
    const currentValue = shares * currentPrice;
    const gain = currentValue - invested;

    totalInvested += invested;
    totalCurrentValue += currentValue;

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
            <h2>${etf.Name}</h2>
            <p><strong>Shares:</strong> ${shares}</p>
            <p><strong>Money Invested:</strong> €${invested.toFixed(2)}</p>
            <p><strong>Average Buy Price:</strong> €${averageBuyPrice.toFixed(
              2
            )}</p>
            <p><strong>Current Price:</strong> €${currentPrice.toFixed(2)}</p>
            <p><strong>Total Current Value:</strong> €${currentValue.toFixed(
              2
            )}</p>
            <p><strong>Gain:</strong> €${gain.toFixed(2)}</p>
        `;
    container.appendChild(card);
  });

  const totalProfit = totalCurrentValue - totalInvested;

  document.getElementById(
    "totalProfit"
  ).textContent = `Total Profit: €${totalProfit.toFixed(2)}`;
  document.getElementById(
    "totalInvested"
  ).textContent = `Total Invested Money: €${totalInvested.toFixed(2)}`;
  document.getElementById(
    "totalCurrentValue"
  ).textContent = `Total Current Value: €${totalCurrentValue.toFixed(2)}`;
}

window.onload = loadData;
