(function() {
  function loadScript(url, cb) {
    var s = document.createElement('script');
    s.src = url; s.onload = cb;
    s.onerror = function() { alert('Không load được: ' + url); };
    document.head.appendChild(s);
  }

  function getAllFilteredData() {
    var table = document.getElementById('bang-chitiet');
    var headers = [];
    table.querySelectorAll('.rt-th').forEach(function(th) {
      var t = th.innerText.trim();
      if (t) headers.push(t);
    });

    var filterMap = window._filterMap || {};
    var filterValues = {};
    Object.keys(filterMap).forEach(function(filterId) {
      var el = document.querySelector('#' + filterId + ' + .selectize-control .selectize-input .item');
      if (el) {
        var val = el.getAttribute('data-value') || el.innerText.trim();
        if (val) filterValues[filterMap[filterId]] = val;
      }
    });

    var allData = window._fullData || [];
    var filtered = allData.filter(function(row) {
      var show = true;
      Object.keys(filterValues).forEach(function(colName) {
        if (row[colName] != filterValues[colName]) show = false;
      });
      return show;
    });

    return { headers: headers, rows: filtered };
  }

  window._taiExcel = function() {
    function doExcel() {
      var d = getAllFilteredData();
      if (!d.rows.length) { alert('Không có dữ liệu!'); return; }
      var wb = XLSX.utils.book_new();
      var ws = XLSX.utils.json_to_sheet(d.rows, { header: d.headers });
      ws['!cols'] = d.headers.map(function(h) { return { wch: Math.max(h.length, 15) }; });
      XLSX.utils.book_append_sheet(wb, ws, 'Danh sach ticket');
      XLSX.writeFile(wb, 'danh_sach_ticket.xlsx');
    }
    window.XLSX ? doExcel() :
      loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', doExcel);
  };

  window._taiPNG = function() {
    function doPNG() {
      var d = getAllFilteredData();
      if (!d.rows.length) { alert('Không có dữ liệu!'); return; }

      var colWidths = window._colWidths || {};

      var tableHtml = '<table style="border-collapse:collapse;font-size:13px;'
        + 'font-family:Arial,sans-serif;table-layout:fixed;">';

      tableHtml += '<colgroup>';
      d.headers.forEach(function(h) {
        var w = colWidths[h] !== undefined ? colWidths[h] : 100;
        tableHtml += '<col style="width:' + w + 'px;">';
      });
      tableHtml += '</colgroup>';

      tableHtml += '<tr>';
      d.headers.forEach(function(h) {
        var w = colWidths[h] !== undefined ? colWidths[h] : 100;
        tableHtml += '<th style="background:#F5DEB3;border:1px solid #ccc;padding:6px 8px;'
          + 'text-align:center;white-space:normal;word-break:break-word;'
          + 'overflow-wrap:break-word;width:' + w + 'px;max-width:' + w + 'px;overflow:hidden;">'
          + h + '</th>';
      });
      tableHtml += '</tr>';

      d.rows.forEach(function(row, idx) {
        var bg = idx % 2 === 0 ? '#ffffff' : '#f2f2f2';
        tableHtml += '<tr style="background:' + bg + ';">';
        d.headers.forEach(function(h) {
          var w = colWidths[h] !== undefined ? colWidths[h] : 100;
          var val = (row[h] != null) ? String(row[h]) : '';
          tableHtml += '<td style="border:1px solid #ccc;padding:5px 8px;'
            + 'white-space:normal;word-break:break-word;overflow-wrap:break-word;'
            + 'vertical-align:top;width:' + w + 'px;max-width:' + w + 'px;overflow:hidden;">'
            + val + '</td>';
        });
        tableHtml += '</tr>';
      });
      tableHtml += '</table>';

      var wrapper = document.createElement('div');
      wrapper.style.cssText = 'position:fixed;top:-99999px;left:-99999px;'
        + 'background:#ffffff;padding:16px;';
      wrapper.innerHTML = tableHtml;
      document.body.appendChild(wrapper);

      setTimeout(function() {
        html2canvas(wrapper, {
          scale: 2, backgroundColor: '#ffffff',
          useCORS: true, allowTaint: true,
          width: wrapper.offsetWidth,
          height: wrapper.offsetHeight
        }).then(function(canvas) {
          document.body.removeChild(wrapper);
          canvas.toBlob(function(blob) {
            if (!blob) { alert('Lỗi tạo ảnh!'); return; }
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url; a.download = 'danh_sach_ticket.png';
            document.body.appendChild(a); a.click();
            setTimeout(function() {
              document.body.removeChild(a);
              URL.revokeObjectURL(url);
            }, 100);
          }, 'image/png');
        }).catch(function(e) {
          if (document.body.contains(wrapper)) document.body.removeChild(wrapper);
          alert('Lỗi PNG: ' + e.message);
        });
      }, 800);
    }
    window.html2canvas ? doPNG() :
      loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js', doPNG);
  };

})();