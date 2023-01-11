/**
 * @NApiVersion 2.x
 * @NScriptType UserEventScript
 * @NModuleScope SameAccount
 */
define(['N/record', 'N/search', 'N/log', 'N/task'],
    /**
     * @param{record} record
     * @param{search} search
     * @param{log} log
     * @param{task} task
     */
    function (record, search, log, task) {

        /**
         * Function definition to be triggered before record is loaded.
         *
         * @param {Object} scriptContext
         * @param {Record} scriptContext.newRecord - New record
         * @param {string} scriptContext.type - Trigger type
         * @param {Form} scriptContext.form - Current form
         * @Since 2015.2
         */
        function beforeLoad(scriptContext) {
            try {
                if (scriptContext.type === scriptContext.UserEventType.VIEW) {
                    var form = scriptContext.form;
                    var request = scriptContext.request;
                    var params = request.parameters;
                    var recordObj = scriptContext.newRecord;
                    var taskID = recordObj.getValue({fieldId: 'custrecord_efx_ms_re_task_id'});
                    var taskStatus = recordObj.getValue({fieldId: 'custrecord_efx_ms_re_status'});
                    var taskDelete = recordObj.getValue({fieldId: 'custrecord_efx_ms_re_task_delete'});
                    var taskStatusDelete = recordObj.getValue({fieldId: 'custrecord_efx_ms_re_status_delete'});
                    log.audit({
                        title: 'task data', details: {
                            taskID: taskID,
                            taskStatus: taskStatus
                        }
                    });
                    if (taskID != "" && taskStatus != 100) {
                        var csvTaskStatus = task.checkStatus({
                            taskId: taskID
                        });
                        log.audit({title: 'status task', details: csvTaskStatus});
                        if (csvTaskStatus.status === task.TaskStatus.COMPLETE) {
                            record.submitFields({
                                type: 'customrecord_efx_ms_rol_env',
                                id: recordObj.id,
                                values: {
                                    'custrecord_efx_ms_re_status': 100
                                }
                            });

                        }
                    } else {
                        if (taskDelete === "") {
                            form.addButton({
                                id: 'custpage_delete_rol',
                                label: 'Borrar Rol',
                                functionName: 'deleteRol'
                            });
                            form.addButton({
                                id: 'custpage_excel_1',
                                label: 'ROL',
                                functionName: 'getExcel1'
                            });
                            form.addButton({
                                id: 'custpage_excel_3',
                                label: 'Consolidado',
                                functionName: 'getExcel3'
                            });
                            form.addButton({
                                id: 'custpage_print_labels',
                                label: 'Imprimir Etiquetas',
                                functionName: 'printLabels'
                            });

                            form.clientScriptModulePath = './efx_ms_del_rol_cs.js';
                        }
                    }
                    if (taskDelete != "" && taskStatusDelete != 100) {
                        var mapStatus = task.checkStatus({taskId: taskDelete});
                        if (mapStatus.status === task.TaskStatus.COMPLETE) {
                            record.submitFields({
                                type: 'customrecord_efx_ms_rol_env',
                                id: recordObj.id,
                                values: {
                                    'custrecord_efx_ms_re_status_delete': 100
                                }
                            });
                        }
                    }
                }
            } catch (e) {
                log.error('Error on beforeLoad', e);
            }

        }

        return {
            beforeLoad: beforeLoad
        };

    });
