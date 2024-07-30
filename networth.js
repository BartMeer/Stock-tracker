function displayDateTime(excelSerial) {
  const date = new Date(Math.round((excelSerial - 25569) * 86400 * 1000));

  date.setMinutes(date.getMinutes() + date.getTimezoneOffset());

  const dateString = date.toDateString();
  const timeString = date.toTimeString().split(" ")[0];
  const formattedDateTime = `${dateString} ${timeString}`;

  document.getElementById("lastUpdated").innerText = formattedDateTime;
}

function loadNetWorthData() {
  fetch("data.xlsx")
    .then((response) => response.arrayBuffer())
    .then((data) => {
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const toDay = worksheet["U3"] ? worksheet["U3"].v : 0;

      displayDateTime(toDay);

      // Define cell ranges for each person
      const row2 = {
        netWorth: worksheet["G2"] ? worksheet["G2"].v : 0,
        bolero: worksheet["H2"] ? worksheet["H2"].v : 0,
        cash: worksheet["I2"] ? worksheet["I2"].v : 0,
        pension: worksheet["J2"] ? worksheet["J2"].v : 0,
      };

      const row3 = {
        netWorth: worksheet["G3"] ? worksheet["G3"].v : 0,
        bolero: worksheet["H3"] ? worksheet["H3"].v : 0,
        cash: worksheet["I3"] ? worksheet["I3"].v : 0,
        kbc: worksheet["J3"] ? worksheet["J3"].v : 0,
      };

      // Calculate totals
      const totalNetWorth =
        parseFloat(row2.netWorth) + parseFloat(row3.netWorth);
      const totalBolero = parseFloat(row2.bolero) + parseFloat(row3.bolero);
      const totalCash = parseFloat(row2.cash) + parseFloat(row3.cash);
      const totalPension = parseFloat(row2.pension);
      const totalKBC = parseFloat(row3.kbc);

      // Update HTML with data for Bart
      document.getElementById(
        "netWorth1"
      ).textContent = `Net Worth: €${parseFloat(row2.netWorth).toFixed(2)}`;
      document.getElementById(
        "allocationBolero1"
      ).textContent = `Bolero Allocation: €${parseFloat(row2.bolero).toFixed(
        2
      )} (${(
        (parseFloat(row2.bolero) / parseFloat(row2.netWorth)) *
        100
      ).toFixed(2)}%)`;
      document.getElementById(
        "allocationCash1"
      ).textContent = `Cash Allocation: €${parseFloat(row2.cash).toFixed(
        2
      )} (${((parseFloat(row2.cash) / parseFloat(row2.netWorth)) * 100).toFixed(
        2
      )}%)`;
      document.getElementById(
        "allocationPension1"
      ).textContent = `Pension Savings Allocation: €${parseFloat(
        row2.pension
      ).toFixed(2)} (${(
        (parseFloat(row2.pension) / parseFloat(row2.netWorth)) *
        100
      ).toFixed(2)}%)`;

      // Update HTML with data for Jolien
      document.getElementById(
        "netWorth2"
      ).textContent = `Net Worth: €${parseFloat(row3.netWorth).toFixed(2)}`;
      document.getElementById(
        "allocationBolero2"
      ).textContent = `Bolero Allocation: €${parseFloat(row3.bolero).toFixed(
        2
      )} (${(
        (parseFloat(row3.bolero) / parseFloat(row3.netWorth)) *
        100
      ).toFixed(2)}%)`;
      document.getElementById(
        "allocationCash2"
      ).textContent = `Cash Allocation: €${parseFloat(row3.cash).toFixed(
        2
      )} (${((parseFloat(row3.cash) / parseFloat(row3.netWorth)) * 100).toFixed(
        2
      )}%)`;
      document.getElementById(
        "allocationKBC"
      ).textContent = `KBC Stocks Allocation: €${parseFloat(row3.kbc).toFixed(
        2
      )} (${((parseFloat(row3.kbc) / parseFloat(row3.netWorth)) * 100).toFixed(
        2
      )}%)`;

      // Update HTML with combined data
      document.getElementById(
        "combinedNetWorth"
      ).textContent = `Net Worth: €${totalNetWorth.toFixed(2)}`;
      document.getElementById(
        "combinedAllocationBolero"
      ).textContent = `Bolero Allocation: €${totalBolero.toFixed(2)} (${(
        (totalBolero / totalNetWorth) *
        100
      ).toFixed(2)}%)`;
      document.getElementById(
        "combinedAllocationCash"
      ).textContent = `Cash Allocation: €${totalCash.toFixed(2)} (${(
        (totalCash / totalNetWorth) *
        100
      ).toFixed(2)}%)`;
      document.getElementById(
        "combinedAllocationPension"
      ).textContent = `Pension Savings Allocation: €${totalPension.toFixed(
        2
      )} (${((totalPension / totalNetWorth) * 100).toFixed(2)}%)`;
      document.getElementById(
        "combinedAllocationKBC"
      ).textContent = `KBC Stocks Allocation: €${totalKBC.toFixed(2)} (${(
        (totalKBC / totalNetWorth) *
        100
      ).toFixed(2)}%)`;
    })
    .catch((error) =>
      console.error("Error fetching or parsing the .xlsx file:", error)
    );
}

window.onload = loadNetWorthData;
