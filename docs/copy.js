function copyTable1() {
  var status = document.getElementById('copy_table_status');
  var table  = document.querySelector('#table1_output table');
  if (!table) {
    status.innerText = 'Error: generate the table first.';
    status.style.color = '#c0392b';
    return;
  }
  var rows = Array.from(table.querySelectorAll('tr'));
  var tsv  = rows.map(function(row) {
    return Array.from(row.querySelectorAll('th, td')).map(function(cell) {
      return cell.innerText.replace(/\n/g, ' ').replace(/\t/g, ' ').trim();
    }).join('\t');
  }).join('\n');

  function onSuccess() {
    status.innerText = 'Copied! Paste into Excel or Google Sheets.';
    status.style.color = '#27ae60';
  }
  function onError(err) {
    status.innerText = 'Error: ' + err;
    status.style.color = '#c0392b';
  }

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsv).then(onSuccess).catch(function(err) {
      try {
        var ta = document.createElement('textarea');
        ta.value = tsv;
        ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.focus(); ta.select();
        var ok = document.execCommand('copy');
        document.body.removeChild(ta);
        ok ? onSuccess() : onError('execCommand returned false');
      } catch(e) { onError(e); }
    });
  } else {
    try {
      var ta = document.createElement('textarea');
      ta.value = tsv;
      ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.focus(); ta.select();
      var ok = document.execCommand('copy');
      document.body.removeChild(ta);
      ok ? onSuccess() : onError('Clipboard API not available in this browser');
    } catch(e) { onError(e); }
  }
}
