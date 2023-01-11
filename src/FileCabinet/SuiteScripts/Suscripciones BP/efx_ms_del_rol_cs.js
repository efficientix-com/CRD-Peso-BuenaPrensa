/**
 * @NApiVersion 2.x
 * @NScriptType ClientScript
 * @NModuleScope SameAccount
 */
define(['N/url', 'N/currentRecord', 'N/ui/message'],
/**
 * @param{url} url
 * @param{currentRecord} currentRecord
 * @param{message} message
 */
function(url, currentRecord, message) {

    var recordForm = currentRecord.get();

    /**
     * Function to be executed after page is initialized.
     *
     * @param {Object} scriptContext
     * @param {Record} scriptContext.currentRecord - Current form record
     * @param {string} scriptContext.mode - The mode in which the record is being accessed (create, copy, or edit)
     *
     * @since 2015.2
     */
    function pageInit(scriptContext) {
        try{
            var recordObj = scriptContext.currentRecord;
            var messageFld = recordObj.getValue({fieldId: 'custpage_fld_error'});
            if (messageFld != '' || messageFld != null) {
                message.create({
                    type: message.Type.ERROR,
                    title: 'Se ha generado un error',
                    message: 'Se genero un error al intentar realizar su peticion, para mas información contacte al administrador'
                }).show();
            }
        } catch (e) {
            console.error(e.message);
        }
    }



    function deleteRol() {
        try{
            var deleteRolMsg = confirm('¿Esta seguro de eliminar el rol?');
            if (deleteRolMsg) {
                var rolID = recordForm.id;
                var output = url.resolveScript({
                    scriptId: 'customscript_efx_ms_del_rol_sl',
                    deploymentId: 'customdeploy_efx_ms_del_rol_sl',
                    params: {
                        'idrol': rolID,
                        'actionsl': 'del'
                    },
                    returnExternalUrl: false,
                });
                window.open(output, '_self');
            }
        } catch (e) {
            console.error(e);
        }
    }

    function getExcelMensual() {
        try {
            var rolID = recordForm.id;
            message.create({
                type: message.Type.INFORMATION,
                title: 'Procesando',
                message: 'Generando Excel Mensual',
                duration: 60000
            }).show()

            var output = url.resolveScript({
                scriptId: 'customscript_efx_ms_del_rol_sl',
                deploymentId: 'customdeploy_efx_ms_del_rol_sl',
                params: {
                    'idrol': rolID,
                    'actionsl': 'e1'
                },
                returnExternalUrl: false,
            });
            window.open(output, '_blank');

        } catch (e) {
            console.error(e);
        }
    }

    function getExcelCAT() {
        try {
            var rolID = recordForm.id;
            alert('Generando Excel CAT Final');

            var output = url.resolveScript({
                scriptId: 'customscript_efx_ms_del_rol_sl',
                deploymentId: 'customdeploy_efx_ms_del_rol_sl',
                params: {
                    'idrol': rolID,
                    'actionsl': 'e2'
                },
                returnExternalUrl: false,
            });
            window.open(output, '_blank');

        } catch (e) {
            console.error(e);
        }
    }

    function getExcelConsolidado() {
        try {
            var rolID = recordForm.id;
            message.create({
                type: message.Type.INFORMATION,
                title: 'Procesando',
                message: 'Generando Excel Consolidado',
                duration: 60000
            }).show()

            var output = url.resolveScript({
                scriptId: 'customscript_efx_ms_del_rol_sl',
                deploymentId: 'customdeploy_efx_ms_del_rol_sl',
                params: {
                    'idrol': rolID,
                    'actionsl': 'e3'
                },
                returnExternalUrl: false,
            });
            window.open(output, '_blank');

        } catch (e) {
            console.error(e);
        }
    }

    function printLabels() {
        try {
            var rolID = recordForm.id;
            message.create({
                type: message.Type.INFORMATION,
                title: 'Procesando',
                message: 'Generando Etiquetas',
                duration: 10000
            }).show()

            var output = url.resolveScript({
                scriptId: 'customscript_efx_ms_con_print_sl',
                deploymentId: 'customdeploy_efx_ms_con_print_sl',
                params: {
                    'idrol': rolID
                },
                returnExternalUrl: false,
            });
            window.open(output, '_blank');
        } catch (e) {
            console.error(e);
        }
    }

    return {
        pageInit: pageInit,
        deleteRol: deleteRol,
        getExcel1: getExcelMensual,
        getExcel2: getExcelCAT,
        getExcel3: getExcelConsolidado,
        printLabels: printLabels
    };

});
