var dataTableArchivo = null;
var TipoArchivo = function () {
      var eventos = function () {
          var modal = $("#modal-medio");

          $("#btn-nuevo").on("click",function () {
                modal.find('.modal-title').text('NUEVO TIPO ARCHIVO');
                $('input[name="action"]').val('add');
                modal.find('i').removeClass().addClass('fas fa-plus');
                $('form')[0].reset();
                 cboTipoProceso(0);

                 modal.modal('show');
          })

          $('#id_aplicaciones').on('change',function () {
                cboTipoProceso();
          });

          $('form').on('submit', function (e) {
                e.preventDefault();
                var valorValidar = '';
                var parameters = new FormData(this);
                if (modal.find('#id_Validar').is(':checked'))
                    valorValidar ='on'
                else
                    valorValidar='off';

                  parameters.delete('Validar');
                  parameters.append('Validar',valorValidar);
                submit_with_ajax(window.location.pathname, 'Notificación', '¿Estas seguro de realizar la siguiente acción?', parameters, function () {
                    modal.modal('hide');
                    dataTableArchivo.ajax.reload();
                });
            });

          $('#data tbody')
          .on('click', 'a[rel="edit"]', function () {
              $('form')[0].reset();
             modal.find('.modal-title').text('EDITAR TIPO DE ARCHIVO');
            modal.find('i').removeClass().addClass('fas fa-edit');
            var tr = dataTableArchivo.cell($(this).closest('td, li')).index();
            var data = dataTableArchivo.row(tr.row).data();

            $('input[name="action"]').val('edit');
            $('input[name="id"]').val(data.id);
            modal.find('#id_aplicaciones').val(data.tipoProcesoId.AplicacionId.id);
            cboTipoProceso(data.tipoProcesoId.id);
            modal.find('#id_Validar').prop("checked",data.Validar);
            modal.find('input[name="NombreTipoCarga"]').val(data.NombreTipoCarga);
            modal.find('input[name="limiteArchivo"]').val(data.limiteArchivo);
            modal.find('input[name="TipoArchivopermitido"]').val(data.TipoArchivopermitido);
            modal.find('input[name="abreviatura"]').val(data.abreviatura);

            modal.modal('show');
        })
          .on('click', 'a[rel="delete"]', function () {
            var tr = dataTableArchivo.cell($(this).closest('td, li')).index();
            var data = dataTableArchivo.row(tr.row).data();
            var parameters = new FormData();
            parameters.append('action', 'delete');
            parameters.append('id', data.id);
            submit_with_ajax(window.location.pathname, 'Notificación', '¿Estas seguro de realizar eliminar el siguiente registro?', parameters, function () {
                dataTableArchivo.ajax.reload();
            });
        });

      }

      var cboTipoProceso =function (id) {
          var idAplicacion =  $("#id_aplicaciones").val();
          var divProceso = $('#id_tipoProcesoId');
          var cboProceso = '';
          if(idAplicacion !=''){
              var parameters = new FormData();
              parameters.append('action', 'getTipoProceso');
              parameters.append('id', idAplicacion);
              divProceso.html('');

              ajaxPOST(window.location.pathname,parameters,function (response) {
                  cboProceso ='<option value="" selected>---------</option>';
                  $.each(response,function (index,value) {
                      cboProceso +='<option value="'+ value.id  +'">'+ value.nombretipoproceso +'</option>';

                  });
                  divProceso.append(cboProceso)
                  if (id>0){
                      divProceso.val(id);
                  }else{
                      divProceso.val('');
                  }

              });
          }else{
                divProceso.html('');
                 cboProceso +='<option value="" selected>---------</option>';
                 divProceso.append(cboProceso)
          }
      }

      var visualDataTableTipoArchivo = function () {
        dataTableArchivo =  $('#data').DataTable({
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
                                console.log(data)
                                return data
                            },
                    },
                    columns: [
                         {"data": "id"},
                        {"data": "tipoProcesoId.AplicacionId.nombreAplicacion"},
                         {"data": "tipoProcesoId.nombretipoproceso"},
                         {"data": "NombreTipoCarga"},
                         {"data": "TipoArchivopermitido"},
                         {"data": "limiteArchivo"},
                         {"data": "abreviatura"},
                    ],
                    columnDefs: [
                        { "bVisible": false, targets: [0] },
                        { "className": "hidden-120", targets: [1], width: '15%' },
                        { "className": "hidden-120", targets: [2], width: '15%' },
                        { "className": "hidden-120", targets: [3], width: '20%' },
                        { "className": "hidden-120", targets: [4], width: '10%' },
                        { "className": "hidden-120", targets: [5], width: '10%' },
                        { "className": "hidden-120", targets: [6], width: '5%' },
                        {targets: [7], class: 'text-center', orderable: false, width:'5%',
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
                visualDataTableTipoArchivo();
        }
    }
}(jQuery);