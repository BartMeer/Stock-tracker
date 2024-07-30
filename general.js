const currencyFormatter = new Intl.NumberFormat('nl-BE', {
  style: 'currency',
  currency: 'EUR',
});

const numberFormatter = new Intl.NumberFormat('nl-BE');

function formatPercentage(part, total) {
  return total === 0 ? '0.00%' : numberFormatter.format((part / total) * 100) + '%';
}

function getCellValue(cell) {
  return cell ? parseFloat(cell.v) || 0 : 0;
}

function updateElement(id, label, value, total = null, showPercentage = false) {
  const element = document.getElementById(id);

  // Format the value and percentage
  const formattedValue = currencyFormatter.format(value);
  let content = `${label}: ${formattedValue}`;

  if (total !== null && showPercentage) {
    const percentage = total === 0 ? 0 : (value / total) * 100;
    const formattedPercentage = numberFormatter.format(percentage.toFixed(2));
    content += ` (${formattedPercentage}%)`;
  }

  // Update the element's text content
  element.textContent = content;
}

function displayDateTime(excelSerial) {
  const date = new Date(Math.round((excelSerial - 25569) * 86400 * 1000));

  date.setMinutes(date.getMinutes() + date.getTimezoneOffset());

  const dateString = date.toLocaleDateString('nl-BE'); // Format date to Belgium locale
  const timeString = date.toLocaleTimeString('nl-BE'); // Format time to Belgium locale
  const formattedDateTime = `${dateString} ${timeString}`;

  document.getElementById('lastUpdated').innerText = formattedDateTime;
}
