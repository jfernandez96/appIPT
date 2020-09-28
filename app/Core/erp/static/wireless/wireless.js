var dataTableCarga=null;
var dataTableCheck =null;
var dataParametro=null;
var CargaArchivo= function () {

    var eventos = function () {
        var modal= $("#modal");
        var modal_desc = $("#modal-Descargar");
        var modal_check =  $("#modal-check");
        $('#btn-nuevo').on("click",function () {
              modal.find(".modal-title").text('CARGA DE ARCHIVOS');
              limpiarForm();
              modal.find('#id_ArcConfigSitio').closest('.form-group').css('display','none');
              modal.find('#id_ArcBaseDF').closest('.form-group').css('display','none');
              modal.find('#id_ArcConfigInicio').closest('.form-group').css('display','none');
              cboTipoProceso();
              modal.modal('show');
        })

       modal.find('#id_aplicaciones').on('change',function () {
            cboTipoProceso();
       });


       modal.find("#id_tipoProcesos").on("change",function () {

           if ($(this).val()!=''){
                   var parameters = new FormData();
                    parameters.append('action', 'getTipo');
                    parameters.append('id', $(this).val());
                    var divFile = $('#divFiles');
                    var file = '';
                    var multiple = '';
                    var accetp='';
                    divFile.html('');
                    ajaxPOST(window.location.pathname,parameters,function (response) {
                        $.each(response,function (index,value) {
                            if(value.limiteArchivo > 1){
                                multiple = 'Multiple';
                            }
                            else{
                                multiple = '';
                            }
                            accetp = value.TipoArchivopermitido.replace('|',',');
                          file ='<div class="form-group"';
                          file +='<label for="'+value.abreviatura+'">'+ value.NombreTipoCarga +'</label>';
                          file +='<input type="file" '+multiple+' accept="'+accetp+'" name="'+value.abreviatura+'" autocomplete="off" class="form-control" required="" id="id_'+value.abreviatura+'"/>'
                          file +='</div>'
                          divFile.append(file)
                        });

                    });
           }

        });

        //$('.select2').select2({
        //        theme: "bootstrap4",
         //       language: 'es'
       // });

        $('#data tbody').on('click', 'a[rel="descargar"]', function () {
            modal_desc.find('.modal-title').text('DESCARGAR ARCHIVOS');
            modal_desc.find('i').removeClass().addClass('fas fa-edit');
            var tr = dataTableCarga.cell($(this).closest('td, li')).index();
            var data = dataTableCarga.row(tr.row).data();

            var parameters = new FormData();
            parameters.append('action','getDetalle');
            parameters.append('id', data.id);
            var link = '';
            var divArchivo= modal_desc.find("#divArchivos");
            divArchivo.html('');
            ajaxPOST(window.location.pathname,parameters,function (response) {
                console.log(response);
                $.each(response,function (index,value) {
                    link ='<li>';
                    link +='<a href="'+  value.ArchivoCheck +'" class="btn-link text-primary"> <text style="font-size:10px; color:#9e9a9a">('+value.tipoArchivoCargaId.NombreTipoCarga+') </text>  '+ value.NombreArchivo +'    <i class="fas fa-download"></i>  </a>';
                    link +='</li>';
                    divArchivo.append(link);
                })

            })
            modal_desc.modal('show');
        })

         .on('click', 'a[rel="delete"]', function () {
            var tr = dataTableCarga.cell($(this).closest('td, li')).index();
            var data = dataTableCarga.row(tr.row).data();
            var parameters = new FormData();
            parameters.append('action', 'delete');
            parameters.append('id', data.id);
            submit_with_ajax(window.location.pathname, 'Notificación', '¿Estas seguro de realizar eliminar el siguiente registro?', parameters, function () {
                dataTableCarga.ajax.reload();
            });
        }).on('click','a[rel="check"]',function () {


             var tr = dataTableCarga.cell($(this).closest('td, li')).index();
             var data = dataTableCarga.row(tr.row).data();

             modal_check.find('.modal-title').text(data.AplicacionId.nombreAplicacion.toUpperCase() +' - '+  data.IdSitio.nombresitio.toUpperCase() +' '+  data.FechaSubida);

             var parameters = new FormData();
            parameters.append('action','getDetalle');
            parameters.append('id', data.id);
            modal_check.find('#id').val(data.id);

            var link = '';
            var divArchivo= modal_check.find("#divArchivos");
            divArchivo.html('');
            ajaxPOST(window.location.pathname,parameters,function (response) {

                var validar = '';
                $.each(response,function (index,value) {
                    if (value.tipoArchivoCargaId.Validar)
                        validar =' (Archivo a validar)';
                    else
                        validar = '';

                    link ='<li class="text-sm-li"> '+ value.NombreArchivo +' <b>'+validar+'</b> </li>';
                    divArchivo.append(link);
                })

            })
            parameters.delete('action');
            parameters.append('action','getParametro');
            ajaxPOST(window.location.pathname,parameters,function (response) {
                dataTableCheck.destroy();
                visualDataTableCheck(response);
                dataParametro = response;
            });
             modal_check.modal('show');
        });

        $("#btnProcesar").on('click',function () {
            var parameters = new FormData();
            parameters.append('action','validarArchivo');
            parameters.append('id', modal_check.find('#id').val());
            parameters.append('dataParametro', JSON.stringify(dataParametro));
            ajaxPOST(window.location.pathname,parameters,function (response) {
                dataParametro = null;
            });

        });

        $('form').on('submit', function (e) {
            e.preventDefault();
            //var parameters = $(this).serializeArray();
            var parameters = new FormData(this);
            submit_with_ajax(window.location.pathname, 'Notificación', '¿Estas seguro de realizar la siguiente acción?', parameters, function () {
               modal.modal('hide');
               dataTableCarga.ajax.reload();
            });
        });
    }

    var limpiarForm = function () {
        var modal= $("#modal");
        modal.find('#id_sitios').val(undefined)
        modal.find('#id_aplicaciones').val('2');
        modal.find('#id_aplicaciones').addClass('isDisabled');
        modal.find('#id_tipoProcesos').val(undefined)
        modal.find('#id_ArcConfigSitio').val('');
        modal.find('#id_ArcBaseDF').val('');
        modal.find('#id_ArcConfigInicio').val('');

    }

    var visualDataTableCarga = function () {
        dataTableCarga=  $('#data').DataTable({
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
                        {"data":"UserId"},
                        {"data": "IdSitio.nombresitio"},
                        {"data": "AplicacionId.nombreAplicacion"},
                        //{"data": "AplicacionId.nombreAplicacion"},
                        {"data": "FechaSubida"},
                    ],
                    columnDefs: [
                            { "className": "hidden-120", targets: [0], width: '10%' },
                            { "className": "hidden-120", targets: [1], width: '40%' },
                            { "className": "hidden-120", targets: [2], width: '20%' },
                            { "className": "hidden-120", targets: [3], width: '20%' },
                        {
                            targets: [4],
                            class: 'text-center',
                            width:'15%',
                            orderable: false,
                            render: function (data, type, row) {
                                var buttons = '<a href="#" rel="descargar" class="btn btn-success btn-xs" title="Descargar archivos"><i class="fas fa-download"></i></a> ';
                                 buttons += '<a href="#" rel="check" type="button" class="btn btn-secondary btn-xs" title="Check"><i class="fas fa-list"></i></a>  ';
                                buttons += '<a href="#" rel="delete" type="button" class="btn btn-danger btn-xs" title="Eliminar"><i class="fas fa-trash-alt"></i></a> ';
                                return buttons;
                            }
                        },
                    ],

                    initComplete: function (settings, json) {

                    }
                })
    }

    var cboTipoProceso = function () {
           var divProceso = $('#id_tipoProcesos');
           var cboProceso = '';

           var idAplicacion = $('#id_aplicaciones').val();

            if (idAplicacion != '' && idAplicacion !=null && idAplicacion){
                    var parameters = new FormData();
                    parameters.append('action', 'getProceso');
                    parameters.append('id',idAplicacion);
                    divProceso.html('');
                    $('#divFiles').html('');
                    ajaxPOST(window.location.pathname,parameters,function (response) {

                        cboProceso ='<option value="" selected>---------</option>';
                        $.each(response,function (index,value) {
                          cboProceso +='<option value="'+ value.id  +'">'+ value.nombretipoproceso +'</option>';

                        });
                        divProceso.append(cboProceso)
                    });
            }else {
                 divProceso.html('');
                 cboProceso +='<option value="" selected>---------</option>';
                 divProceso.append(cboProceso)
                  $('#divFiles').html('');
            }
    }

    var visualDataTableCheck =function (jsonData) {
        dataTableCheck = $("#tableCheck").DataTable({
             "bFilter": true,
             "bPaginate": true,
             "ordering": true,
             destroy: true,
             "data": jsonData,
            "columns": [
                { "data": "mo" },
                { "data": "parametro" },
                { "data": "valor" },
                {"data" : function (obj) {
                    return '<i class="fa fa-stopwatch text-warning" title="sin procesar"></i>'
                    }
                }
                ],
              "aoColumnDefs": [
                { "className": "center hidden-120", "aTargets": [0], "width": "30%" },
                { "className": "center hidden-120", "aTargets": [1], "width": "40%" },
                { "className": "center hidden-120", "aTargets": [2], "width": "20%" },
                { "className": "center hidden-120", "aTargets": [3], "width": "10%" },
              ],
              "order": [[0, "desc"]],
         });
    }
    return{
        init:function () {
            eventos();
            visualDataTableCheck()
            visualDataTableCarga();
        }
    }
}(jQuery);

