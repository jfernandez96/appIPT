var dataTableProceso=null;
var TipoProceso = function () {
      var eventos = function () {
          var modal = $("#modal-medio");

          $("#btn-nuevo").on("click",function () {
                modal.find('.modal-title').text('NUEVO TIPO DE PROCESO');
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
                    dataTableProceso.ajax.reload();
                });
            });

          $('#data tbody')
          .on('click', 'a[rel="edit"]', function () {
             modal.find('.modal-title').text('EDITAR TIPO DE PROCESO');
            modal.find('i').removeClass().addClass('fas fa-edit');
            var tr = dataTableProceso.cell($(this).closest('td, li')).index();
            var data = dataTableProceso.row(tr.row).data();
            console.log(data.AplicacionId.id)
            $('input[name="action"]').val('edit');
            $('input[name="id"]').val(data.id);
            modal.find('#id_AplicacionId').val(data.AplicacionId.id);
            modal.find('input[name="nombretipoproceso"]').val(data.nombretipoproceso);
            modal.modal('show');
        })
        .on('click', 'a[rel="delete"]', function () {
            var tr = dataTableProceso.cell($(this).closest('td, li')).index();
            var data = dataTableProceso.row(tr.row).data();
            var parameters = new FormData();
            parameters.append('action', 'delete');
            parameters.append('id', data.id);
            submit_with_ajax(window.location.pathname, 'Notificación', '¿Estas seguro de realizar eliminar el siguiente registro?', parameters, function () {
                dataTableProceso.ajax.reload();
            });
        });

      }

      var visualDataTableTipo = function () {
        dataTableProceso =  $('#data').DataTable({
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
                         {"data": "AplicacionId.nombreAplicacion"},
                         {"data": "nombretipoproceso"},
                    ],
                    columnDefs: [
                        { "bVisible": false, targets: [0] },
                        { "className": "hidden-120", targets: [1], width: '45%' },
                        { "className": "hidden-120", targets: [2], width: '45%' },
                        {targets: [3], class: 'text-center', orderable: false, width:'5%',
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
                visualDataTableTipo();
        }
    }
}(jQuery);