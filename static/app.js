/* Tiện ích dùng chung: DataTables (tiếng Việt, tìm không dấu), chọn nhiều dòng, gọi API, toast */
(function () {
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D');
  window.norm = norm;
  // Tìm kiếm không phân biệt dấu tiếng Việt (cả ô tìm lẫn dữ liệu đều bỏ dấu)
  $.fn.dataTable.ext.type.search.string = d => norm(d);
  $.fn.dataTable.ext.type.search.html = d => norm(String(d == null ? '' : d).replace(/<[^>]*>/g, ' '));

  window.toast = function (msg, kind) {
    const box = document.getElementById('toast'), d = document.createElement('div');
    d.textContent = msg; if (kind) d.className = kind; box.appendChild(d);
    setTimeout(() => d.remove(), 3500);
  };

  window.postJSON = async function (url, data) {
    const r = await fetch(url, {method: 'POST', headers: {'Content-Type': 'application/json', 'X-CSRF': window.CSRF},
      body: JSON.stringify(data)});
    let j = {}; try { j = await r.json(); } catch (e) {}
    if (!r.ok) throw new Error(j.error || ('Lỗi ' + r.status));
    return j;
  };

  const LANG = {search: 'Tìm:', lengthMenu: 'Hiện _MENU_ dòng', info: '_START_–_END_ / _TOTAL_ dòng',
    infoEmpty: 'Không có dữ liệu', infoFiltered: '(lọc từ _MAX_ dòng)', zeroRecords: 'Không tìm thấy dòng phù hợp',
    emptyTable: 'Chưa có dữ liệu', paginate: {first: '«', last: '»', next: '›', previous: '‹'}};

  /* Khởi tạo DataTable: ô tìm #q, cột STT (th.stt) tự đánh lại, cột không sắp xếp (th.nosort) */
  window.initTable = function (sel, opts) {
    const t = $(sel).DataTable(Object.assign({
      language: LANG, order: [], pageLength: 25, autoWidth: false,
      lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'Tất cả']],
      dom: 'lrt<"dt-foot"ip>',
      columnDefs: [{orderable: false, searchable: false, targets: 'nosort'}]
    }, opts || {}));
    const renum = () => {
      const start = t.page.info().start;
      t.column('.stt', {page: 'current'}).nodes().each((c, i) => { c.textContent = start + i + 1; });
    };
    t.on('draw.dt', renum); renum();
    $('#q').on('input', function () { t.search(norm(this.value)).draw(); });
    return t;
  };

  /* Lọc cột bằng <select>: điền sẵn giá trị duy nhất của cột; trả về hàm áp dụng lại */
  window.colFilter = function (t, selEl, idx) {
    const el = $(selEl), vals = t.column(idx).data().unique().sort().toArray().filter(v => String(v).trim());
    vals.forEach(v => el.append($('<option>').text(String(v).replace(/<[^>]*>/g, ''))));
    el.on('change', function () {
      const v = this.value;
      t.column(idx).search(v ? '^' + $.fn.dataTable.util.escapeRegex(norm(v)) + '$' : '', true, false).draw();
    });
  };

  /* Chọn nhiều dòng (giữ trạng thái qua các trang) + thanh thao tác hàng loạt */
  window.bulkSelect = function (t) {
    const all = $('#chkAll'), bar = $('#bulkbar');
    const sel = () => t.$('input.chk:checked').map((i, e) => e.value).get();
    const upd = () => {
      const n = sel().length; bar.toggleClass('show', n > 0); $('#bulkcount').text(n);
      const tot = t.rows({search: 'applied'}).nodes().to$().find('input.chk');
      all.prop('checked', n > 0 && tot.length > 0 && tot.filter(':checked').length === tot.length);
    };
    $(t.table().body()).on('change', 'input.chk', upd);
    all.on('change', function () {
      t.rows({search: 'applied'}).nodes().to$().find('input.chk').prop('checked', this.checked); upd();
    });
    $('#bulkclear').on('click', () => { t.$('input.chk').prop('checked', false); upd(); });
    t.on('draw.dt', upd);
    return {sel, update: upd};
  };
})();
