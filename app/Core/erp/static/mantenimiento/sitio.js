var dataTableSitio=null;
var Sitio = function () {
      var eventos = function () {
          var modal = $("#modal-medio");

          $("#btn-nuevo").on("click",function () {
                modal.find('.modal-title').text('NUEVO SITIO');
                $('input[name="action"]').val('add');
                console.log(modal.find('i'));
                modal.find('i').removeClass().addClass('fas fa-plus');
                $('form')[0].reset();
               modal.modal('show');
          })

          $('form').on('submit', function (e) {
                e.preventDefault();
                //var parameters = $(this).serializeArray();
                var parameters = new FormData(this);
                submit_with_ajax(window.location.pathname, 'Notificación', '¿Estas seguro de realizar la siguiente acción?', parameters, function () {
                    modal.modal('hide');
                    dataTableSitio.ajax.reload();
                });
            });

          $('#data tbody')
          .on('click', 'a[rel="edit"]', function () {
             modal.find('.modal-title').text('EDITAR SITIO');
            modal.find('i').removeClass().addClass('fas fa-edit');
            var tr = dataTableSitio.cell($(this).closest('td, li')).index();
            var data = dataTableSitio.row(tr.row).data();
            $('input[name="action"]').val('edit');
            $('input[name="id"]').val(data.id);
            $('input[name="nombresitio"]').val(data.nombresitio);
            modal.modal('show');
        })
        .on('click', 'a[rel="delete"]', function () {
            var tr = dataTableSitio.cell($(this).closest('td, li')).index();
            var data = dataTableSitio.row(tr.row).data();
            var parameters = new FormData();
            parameters.append('action', 'delete');
            parameters.append('id', data.id);
            submit_with_ajax(window.location.pathname, 'Notificación', '¿Estas seguro de realizar eliminar el siguiente registro?', parameters, function () {
                dataTableSitio.ajax.reload();
            });
        });

      }

      var visualDataTableSitio = function () {
        dataTableSitio =  $('#data').DataTable({
                    responsive: true,
                    autoWidth: false,
                    destroy: true,
                    deferRender: true,
                    ajax: {
                        url: window.location.pathname,
                        type: 'POST',
                        data: {
                            'action': 'searchdata'
                        },
                        dataSrc: "",
                        dataFilter: function (data) {
                                return data
                            },
                    },
                    columns: [
                         {"data": "id"},
                        {"data": "nombresitio"},
                    ],
                    columnDefs: [
                        { "bVisible": false, targets: [0] },
                        { "className": "hidden-120", targets: [1], width: '95%' },
                        {targets: [2], class: 'text-center', orderable: false, width:'5%',
                            render: function (data, type, row) {
                                 var buttons = '<a href="#" rel="edit" class="btn btn-warning btn-xs btnEdit"><i class="fas fa-edit"></i></a> ';
                                    buttons += '<a href="#" rel="delete" class="btn btn-danger btn-xs"><i class="fas fa-trash-alt"></i></a>';
                                 return buttons;
                            }
                        },
                    ],
                    initComplete: function (settings, json) {

                    }
                })
    }

    return{
        init:function () {
                eventos();
                visualDataTableSitio();
        }
    }
}(jQuery);