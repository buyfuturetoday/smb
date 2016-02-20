/* Project specific JavaScript OM library */
/* Version: 15.0.4225.1000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

OSF.ClientMode={
	ReadWrite: 0,
	ReadOnly: 1
}
OSF.DDA.RichInitializationReason={
	1: Microsoft.Office.WebExtension.InitializationReason.Inserted,
	2: Microsoft.Office.WebExtension.InitializationReason.DocumentOpened
};
OSF.DDA.RichClientSettingsManager={
	read: function OSF_DDA_RichClientSettingsManager$Read(onCalling, onReceiving) {
		var keys=[];
		var values=[];
		if (onCalling) {
			onCalling();
		}
		window.external.GetContext().GetSettings().Read(keys, values);
		if (onReceiving) {
			onReceiving();
		}
		var serializedSettings={};
		for (var index=0; index < keys.length; index++) {
			serializedSettings[keys[index]]=values[index];
		}
		return serializedSettings;
	},
	write: function OSF_DDA_RichClientSettingsManager$Write(serializedSettings, overwriteIfStale, onCalling, onReceiving) {
		var keys=[];
		var values=[];
		for (var key in serializedSettings) {
			keys.push(key);
			values.push(serializedSettings[key]);
		}
		if (onCalling) {
			onCalling();
		}
		window.external.GetContext().GetSettings().Write(keys, values);
		if (onReceiving) {
			onReceiving();
		}
	}
};
OSF.DDA.DispIdHost.getRichClientDelegateMethods=function (actionId) {
	var delegateMethods={};
	delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;
	delegateMethods[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;
	function getSettingsExecuteMethod(hostDelegateMethod) {
		return function (args) {
			var status, response;
			try {
				response=hostDelegateMethod(args.hostCallArgs, args.onCalling, args.onReceiving);
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess;
			} catch (ex) {
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				response={ name : Strings.OfficeOM.L_InternalError, message : ex };
			}
			if (args.onComplete) {
				args.onComplete(status, response);
			}
		};
	}
	function readSerializedSettings(hostCallArgs, onCalling, onReceiving) {
		return OSF.DDA.RichClientSettingsManager.read(onCalling, onReceiving);
	}
	function writeSerializedSettings(hostCallArgs, onCalling, onReceiving) {
		return OSF.DDA.RichClientSettingsManager.write(
			hostCallArgs[OSF.DDA.SettingsManager.SerializedSettings],
			hostCallArgs[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale],
			onCalling,
			onReceiving
		);
	}
	switch (actionId) {
		case OSF.DDA.AsyncMethodNames.RefreshAsync.id:
			delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=getSettingsExecuteMethod(readSerializedSettings);
			break;
		case OSF.DDA.AsyncMethodNames.SaveAsync.id:
			delegateMethods[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=getSettingsExecuteMethod(writeSerializedSettings);
			break;
		default:
			break;
	}
	return delegateMethods;
}
OSF.DDA.CustomXmlParts=function OSF_DDA_CustomXmlParts() {
	this._eventDispatches=[];
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.AddDataPartAsync,
		am.GetDataPartByIdAsync,
		am.GetDataPartsByNameSpaceAsync
	]);
};
OSF.DDA.CustomXmlPart=function OSF_DDA_CustomXmlPart(customXmlParts, id, builtIn) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"builtIn": {
			value: builtIn
		},
		"id": {
			value: id
		},
		"namespaceManager": {
			value: new OSF.DDA.CustomXmlPrefixMappings(id)
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.DeleteDataPartAsync,
		am.GetPartNodesAsync,
		am.GetPartXmlAsync
	]);
	var customXmlPartEventDispatches=customXmlParts._eventDispatches;
	var dispatch=customXmlPartEventDispatches[id];
	if (!dispatch) {
		var et=Microsoft.Office.WebExtension.EventType;
		dispatch=new OSF.EventDispatch([
			et.DataNodeDeleted,
			et.DataNodeInserted,
			et.DataNodeReplaced
		]);
		customXmlPartEventDispatches[id]=dispatch;
	}
	OSF.DDA.DispIdHost.addEventSupport(this, dispatch);
};
OSF.DDA.CustomXmlPrefixMappings=function OSF_DDA_CustomXmlPrefixMappings(partId) {
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this,
		[
			am.AddDataPartNamespaceAsync,
			am.GetDataPartNamespaceAsync,
			am.GetDataPartPrefixAsync
		],
		partId
	);
};
OSF.DDA.CustomXmlNode=function OSF_DDA_CustomXmlNode(handle, nodeType, ns, baseName) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"baseName": {
			value: baseName
		},
		"namespaceUri": {
			value: ns
		},
		"nodeType": {
			value: nodeType
		}
	});
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(
		this,
		[
			am.GetRelativeNodesAsync,
			am.GetNodeValueAsync,
			am.GetNodeXmlAsync,
			am.SetNodeValueAsync,
			am.SetNodeXmlAsync
		],
		handle
	);
};
OSF.DDA.NodeInsertedEventArgs=function OSF_DDA_NodeInsertedEventArgs(newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeInserted
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeReplacedEventArgs=function OSF_DDA_NodeReplacedEventArgs(oldNode, newNode, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeReplaced
		},
		"oldNode": {
			value: oldNode
		},
		"newNode": {
			value: newNode
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.DDA.NodeDeletedEventArgs=function OSF_DDA_NodeDeletedEventArgs(oldNode, oldNextSibling, inUndoRedo) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.DataNodeDeleted
		},
		"oldNode": {
			value: oldNode
		},
		"oldNextSibling": {
			value: oldNextSibling
		},
		"inUndoRedo": {
			value: inUndoRedo
		}
	});
};
OSF.OUtil.setNamespace("SafeArray", OSF.DDA);
OSF.DDA.SafeArray.Response={
	Status: 0,
	Payload: 1
};
OSF.DDA.SafeArray.UniqueArguments={
	Offset: "offset",
	Run: "run",
	BindingSpecificData: "bindingSpecificData",
	MergedCellGuid: "{66e7831f-81b2-42e2-823c-89e872d541b3}"
};
OSF.OUtil.setNamespace("Delegate", OSF.DDA.SafeArray);
OSF.DDA.SafeArray.Delegate.SpecialProcessor=function OSF_DDA_SafeArray_Delegate_SpecialProcessor() {
	function _2DVBArrayToJaggedArray(vbArr) {
		var ret;
		try {
			var rows=vbArr.ubound(1);
			var cols=vbArr.ubound(2);
			vbArr=vbArr.toArray();
			if (rows==1 && cols==1) {
				ret=[vbArr];
			} else {
				ret=[];
				for (var row=0; row < rows; row++) {
					var rowArr=[];
					for (var col=0; col < cols; col++) {
						var datum=vbArr[row * cols+col];
						if (datum !=OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid) {
							rowArr.push(datum);
						}
					}
					if (rowArr.length > 0) {
						ret.push(rowArr);
					}
				}
			}
		} catch (ex) {
		}
		return ret;
	}
	var complexTypes=[
		OSF.DDA.PropertyDescriptors.BindingProperties,
		OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
		OSF.DDA.SafeArray.UniqueArguments.Offset,
		OSF.DDA.SafeArray.UniqueArguments.Run,
		OSF.DDA.PropertyDescriptors.Subset,
		OSF.DDA.PropertyDescriptors.DataPartProperties,
		OSF.DDA.PropertyDescriptors.DataNodeProperties,
		OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
		OSF.DDA.EventDescriptors.DataNodeInsertedEvent,
		OSF.DDA.EventDescriptors.DataNodeReplacedEvent,
		OSF.DDA.EventDescriptors.DataNodeDeletedEvent,
		OSF.DDA.DataNodeEventProperties.OldNode,
		OSF.DDA.DataNodeEventProperties.NewNode,
		OSF.DDA.DataNodeEventProperties.NextSiblingNode
	];
	var dynamicTypes={};
	dynamicTypes[Microsoft.Office.WebExtension.Parameters.Data]=(function () {
		var tableRows=0;
		var tableHeaders=1;
		return {
			toHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$toHost(data) {
				if (typeof data !="string" && data[OSF.DDA.TableDataProperties.TableRows] !==undefined) {
					var tableData=[];
					tableData[tableRows]=data[OSF.DDA.TableDataProperties.TableRows];
					tableData[tableHeaders]=data[OSF.DDA.TableDataProperties.TableHeaders];
					data=tableData;
				}
				return data;
			},
			fromHost: function OSF_DDA_SafeArray_Delegate_SpecialProcessor_Data$fromHost(hostArgs) {
				var ret;
				if (hostArgs.toArray) {
					var dimensions=hostArgs.dimensions();
					if(dimensions===2) {
						ret=_2DVBArrayToJaggedArray(hostArgs);
					} else {
						var array=hostArgs.toArray();
						if(array.length===2 &&  ((array[0] !=null && array[0].toArray) || (array[1] !=null && array[1].toArray))) {
							ret={};
							ret[OSF.DDA.TableDataProperties.TableRows]=_2DVBArrayToJaggedArray(array[tableRows]);
							ret[OSF.DDA.TableDataProperties.TableHeaders]=_2DVBArrayToJaggedArray(array[tableHeaders]);
						} else {
							ret=array;
						}
					}
				} else {
					ret=hostArgs;
				}
				return ret;
			}
		}
	})();
	OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(this, complexTypes, dynamicTypes);
	this.pack=function OSF_DDA_SafeArray_Delegate_SpecialProcessor$pack(param, arg) {
		var value;
		if (this.isDynamicType(param)) {
			value=dynamicTypes[param].toHost(arg);
		} else {
			value=arg;
		}
		return value;
	};
	this.unpack=function OSF_DDA_SafeArray_Delegate_SpecialProcessor$unpack(param, arg) {
		var value;
		if (this.isComplexType(param) || OSF.DDA.ListType.isListType(param)) {
			try {
				value=arg.toArray();
			} catch (ex) {
				value=arg || {};
			}
		} else if (this.isDynamicType(param)) {
			value=dynamicTypes[param].fromHost(arg);
		} else {
			value=arg;
		}
		return value;
	};
}
OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor, OSF.DDA.SpecialProcessor);
OSF.DDA.SafeArray.Delegate.ParameterMap=(function () {
	var parameterMap=new OSF.DDA.HostParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor());
	var ns;
	var self=parameterMap.self;
	function createObject(properties) {
		var obj=null;
		if (properties) {
			obj={};
			var len=properties.length;
			for (var i=0; i < len; i++) {
				obj[properties[i].name]=properties[i].value;
			}
		}
		return obj;
	}
	function define(definition) {
		var args={};
		var toHost=createObject(definition.toHost);
		if (definition.invertible) {
			args.map=toHost;
		}
		else if (definition.canonical) {
			args.toHost=args.fromHost=toHost;
		}
		else {
			args.toHost=toHost;
			args.fromHost=createObject(definition.fromHost);
		}
		parameterMap.setMapping(definition.type, args);
	}
	ns=OSF.DDA.BindingProperties;
	define({
		type: OSF.DDA.PropertyDescriptors.BindingProperties,
		fromHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.Type, value: 1 },
			{ name: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData, value: 2 }
		]
	});
	define({
		type: OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,
		fromHost: [
			{ name: ns.RowCount, value: 0 },
			{ name: ns.ColumnCount, value: 1 },
			{ name: ns.HasHeaders, value: 2 }
		]
	});
	ns=OSF.DDA.SafeArray.UniqueArguments;
	define({
		type: OSF.DDA.PropertyDescriptors.Subset,
		toHost: [
			{ name: ns.Offset, value: 0 },
			{ name: ns.Run, value: 1 }
		],
		canonical: true
	});
	ns=Microsoft.Office.WebExtension.Parameters;
	define({
		type: OSF.DDA.SafeArray.UniqueArguments.Offset,
		toHost: [
			{ name: ns.StartRow, value: 0 },
			{ name: ns.StartColumn, value: 1 }
		],
		canonical: true
	});
	define({
		type: OSF.DDA.SafeArray.UniqueArguments.Run,
		toHost: [
			{ name: ns.RowCount, value: 0 },
			{ name: ns.ColumnCount, value: 1 }
		],
		canonical: true
	});
	ns=OSF.DDA.DataPartProperties;
	define({
		type: OSF.DDA.PropertyDescriptors.DataPartProperties,
		fromHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.BuiltIn, value: 1 }
		]
	});
	ns=OSF.DDA.DataNodeProperties;
	define({
		type: OSF.DDA.PropertyDescriptors.DataNodeProperties,
		fromHost: [
			{ name: ns.Handle, value: 0 },
			{ name: ns.BaseName, value: 1 },
			{ name: ns.NamespaceUri, value: 2 },
			{ name: ns.NodeType, value: 3 }
		]
	});
	define({
		type: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: 0 },
			{ name: OSF.DDA.PropertyDescriptors.Subset, value: 1 }
		]
	});
	ns=OSF.DDA.DataNodeEventProperties;
	define({
		type: OSF.DDA.EventDescriptors.DataNodeInsertedEvent,
		fromHost: [
			{ name: ns.InUndoRedo, value: 0 },
			{ name: ns.NewNode, value: 1 }
		]
	});
	define({
		type: OSF.DDA.EventDescriptors.DataNodeReplacedEvent,
		fromHost: [
			{ name: ns.InUndoRedo, value: 0 },
			{ name: ns.OldNode, value: 1 },
			{ name: ns.NewNode, value: 2 }
		]
	});
	define({
		type: OSF.DDA.EventDescriptors.DataNodeDeletedEvent,
		fromHost: [
			{ name: ns.InUndoRedo, value: 0 },
			{ name: ns.OldNode, value: 1 },
			{ name: ns.NextSiblingNode, value: 2 }
		]
	});
	define({
		type: ns.OldNode,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: self }
		]
	});
	define({
		type: ns.NewNode,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: self }
		]
	});
	define({
		type: ns.NextSiblingNode,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataNodeProperties, value: self }
		]
	});
	ns=Microsoft.Office.WebExtension.AsyncResultStatus;
	define({
		type: OSF.DDA.PropertyDescriptors.AsyncResultStatus,
		fromHost: [
			{ name: ns.Succeeded, value: 0 },
			{ name: ns.Failed, value: 1 }
		]
	});
	ns=Microsoft.Office.WebExtension.CoercionType;
	define({
		type: Microsoft.Office.WebExtension.Parameters.CoercionType,
		toHost: [
			{ name: ns.Text, value: 0 },
			{ name: ns.Matrix, value: 1 },
			{ name: ns.Table, value: 2 },
			{ name: ns.Html, value: 3 },
			{ name: ns.Ooxml, value: 4 }
		]
	});
	ns=Microsoft.Office.WebExtension.FileType;
	if (ns) {
		define({
			type: Microsoft.Office.WebExtension.Parameters.FileType,
			toHost: [
			{ name: ns.Text, value: 0 },
			{ name: ns.Base64, value: 5 }
		]
		});
	}
	ns=Microsoft.Office.WebExtension.BindingType;
	define({
		type: Microsoft.Office.WebExtension.Parameters.BindingType,
		toHost: [
			{ name: ns.Text, value: 0 },
			{ name: ns.Matrix, value: 1 },
			{ name: ns.Table, value: 2 }
		],
		invertible: true
	});
	ns=Microsoft.Office.WebExtension.ValueFormat;
	define({
		type: Microsoft.Office.WebExtension.Parameters.ValueFormat,
		toHost: [
			{ name: ns.Unformatted, value: 0 },
			{ name: ns.Formatted, value: 1 }
		]
	});
	ns=Microsoft.Office.WebExtension.FilterType;
	define({
		type: Microsoft.Office.WebExtension.Parameters.FilterType,
		toHost: [
			{ name: ns.All, value: 0 },
			{ name: ns.OnlyVisible, value: 1 }
		]
	});
	ns=Microsoft.Office.WebExtension.Parameters;
	var cns=OSF.DDA.MethodDispId;
	define({
		type: cns.dispidGetSelectedDataMethod,
		fromHost: [
			{ name: ns.Data, value: self }
		],
		toHost: [
			{ name: ns.CoercionType, value: 0 },
			{ name: ns.ValueFormat, value: 1 },
			{ name: ns.FilterType, value: 2 }
		]
	});
	define({
		type: cns.dispidSetSelectedDataMethod,
		toHost: [
			{ name: ns.CoercionType, value: 0 },
			{ name: ns.Data, value: 1 }
		]
	});
	define({
		type: cns.dispidGetWholeDocumentMethod,
		toHost: [{ name: ns.FileType, value: 0}],
		fromHost: [{ name: ns.Data, value: self}]
	});
	define({
		type: cns.dispidAddBindingFromSelectionMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.BindingType, value: 1 }
		]
	});
	define({
		type: cns.dispidAddBindingFromPromptMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.BindingType, value: 1 },
			{ name: ns.PromptText, value: 2 }
		]
	});
	define({
		type: cns.dispidAddBindingFromNamedItemMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.ItemName, value: 0 },
			{ name: ns.Id, value: 1 },
			{ name: ns.BindingType, value: 2 },
			{ name: ns.FailOnCollision, value: 3 }
		]
	});
	define({
		type: cns.dispidReleaseBindingMethod,
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetBindingMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetAllBindingsMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.BindingList, value: self }
		]
	});
	define({
		type: cns.dispidGetBindingDataMethod,
		fromHost: [
			{ name: ns.Data, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.CoercionType, value: 1 },
			{ name: ns.ValueFormat, value: 2 },
			{ name: ns.FilterType, value: 3 },
			{ name: OSF.DDA.PropertyDescriptors.Subset, value: 4 }
		]
	});
	define({
		type: cns.dispidSetBindingDataMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.CoercionType, value: 1 },
			{ name: ns.Data, value: 2 },
			{ name: OSF.DDA.SafeArray.UniqueArguments.Offset, value: 3 }
		]
	});
	define({
		type: cns.dispidAddRowsMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.Data, value: 1 }
		]
	});
	define({
		type: cns.dispidAddColumnsMethod,
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.Data, value: 1 }
		]
	});
	define({
		type: cns.dispidClearAllRowsMethod,
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidLoadSettingsMethod,
		fromHost: [
			{ name: OSF.DDA.SettingsManager.SerializedSettings, value: self }
		]
	});
	define({
		type: cns.dispidSaveSettingsMethod,
		toHost: [
			{ name: OSF.DDA.SettingsManager.SerializedSettings, value: OSF.DDA.SettingsManager.SerializedSettings },
			{ name: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale, value: Microsoft.Office.WebExtension.Parameters.OverwriteIfStale }
		]
	});
	define({
		type: cns.dispidAddDataPartMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataPartProperties, value: self }
		],
		toHost: [
			{ name: ns.Xml, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartByIdMethod,
		fromHost: [
			{ name: OSF.DDA.PropertyDescriptors.DataPartProperties, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartsByNamespaceMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.DataPartList, value: self }
		],
		toHost: [
			{ name: ns.Namespace, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartXmlMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataPartNodesMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.DataNodeList, value: self }
		],
		toHost: [
			{ name: ns.Id, value: 0 },
			{ name: ns.XPath, value: 1 }
		]
	});
	define({
		type: cns.dispidDeleteDataPartMethod,
		toHost: [
			{ name: ns.Id, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataNodeValueMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataNodeXmlMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 }
		]
	});
	define({
		type: cns.dispidGetDataNodesMethod,
		fromHost: [
			{ name: OSF.DDA.ListDescriptors.DataNodeList, value: self }
		],
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
			{ name: ns.XPath, value: 1 }
		]
	});
	define({
		type: cns.dispidSetDataNodeValueMethod,
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
			{ name: ns.Data, value: 1 }
		]
	});
	define({
		type: cns.dispidSetDataNodeXmlMethod,
		toHost: [
			{ name: OSF.DDA.DataNodeProperties.Handle, value: 0 },
			{ name: ns.Xml, value: 1 }
		]
	});
	define({
		type: cns.dispidAddDataNamespaceMethod,
		toHost: [
			{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
			{ name: ns.Prefix, value: 1 },
			{ name: ns.Namespace, value: 2 }
		]
	});
	define({
		type: cns.dispidGetDataUriByPrefixMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
			{ name: ns.Prefix, value: 1 }
		]
	});
	define({
		type: cns.dispidGetDataPrefixByUriMethod,
		fromHost: [
			{ name: ns.Data, value: self}
		],
		toHost: [
			{ name: OSF.DDA.DataPartProperties.Id, value: 0 },
			{ name: ns.Namespace, value: 1 }
		]
	});
	define({
		type: cns.dispidGetSelectedTaskMethod,
		fromHost: [
			{ name: ns.TaskId, value: self }
		]
	});
	define({
		type: cns.dispidGetTaskMethod,
		fromHost: [
			{ name: "taskName", value: 0 },
			{ name: "wssTaskId", value: 1 },
			{ name: "resourceNames", value: 2 }
		],
		toHost: [
			{ name: ns.TaskId, value: 0 }
		]
	});
	define({
		type: cns.dispidGetTaskFieldMethod,
		fromHost: [
			{ name: ns.FieldValue, value: self }
		],
		toHost: [
			{ name: ns.TaskId, value: 0 },
			{ name: ns.FieldId, value: 1 },
			{ name: ns.GetRawValue, value: 2 }
		]
	});
	define({
		type: cns.dispidGetWSSUrlMethod,
		fromHost: [
			{ name: ns.ServerUrl, value: 0 },
			{ name: ns.ListName, value: 1 }
		]
	});
	define({
		type: cns.dispidGetSelectedResourceMethod,
		fromHost: [
			{ name: ns.ResourceId, value: self }
		]
	});
	define({
		type: cns.dispidGetResourceFieldMethod,
		fromHost: [
			{ name: ns.FieldValue, value: self }
		],
		toHost: [
			{ name: ns.ResourceId, value: 0 },
			{ name: ns.FieldId, value: 1 },
			{ name: ns.GetRawValue, value: 2 }
		]
	});
	define({
		type: cns.dispidGetProjectFieldMethod,
		fromHost: [
			{ name: ns.FieldValue, value: self }
		],
		toHost: [
			{ name: ns.FieldId, value: 0 },
			{ name: ns.GetRawValue, value: 1 }
		]
	});
	define({
		type: cns.dispidGetSelectedViewMethod,
		fromHost: [
			{ name: ns.ViewType, value: 0 },
			{ name: ns.ViewName, value: 1 }
		]
	});
	cns=OSF.DDA.EventDispId
	define({ type: cns.dispidDocumentSelectionChangedEvent });
	define({
		type: cns.dispidBindingSelectionChangedEvent,
		fromHost: [
			{name: OSF.DDA.EventDescriptors.BindingSelectionChangedEvent, value: self}
		]
	});
	define({
		type: cns.dispidBindingDataChangedEvent,
		fromHost: [{ name: OSF.DDA.PropertyDescriptors.BindingProperties, value: self}]
	});
	define({ type: cns.dispidSettingsChangedEvent });
	define({
		type: cns.dispidDataNodeAddedEvent,
		fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeInsertedEvent, value: self}]
	});
	define({
		type: cns.dispidDataNodeReplacedEvent,
		fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeReplacedEvent, value: self}]
	});
	define({
		type: cns.dispidDataNodeDeletedEvent,
		fromHost: [{ name: OSF.DDA.EventDescriptors.DataNodeDeletedEvent, value: self}]
	});
	define({ type: cns.dispidTaskSelectionChangedEvent });
	define({ type: cns.dispidResourceSelectionChangedEvent });
	define({ type: cns.dispidViewSelectionChangedEvent });
	return parameterMap;
})();
OSF.DDA.SafeArray.Delegate._onException=function OSF_DDA_SafeArray_Delegate$OnException(ex, args) {
	var status;
	var number=ex.number;
	if (number) {
		switch (number) {
			case -2146828218:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability;
				break;
			case -2146827850:
			default:
				status=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;
				break;
		}
	}
	if (args.onComplete) {
		args.onComplete(status || OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError);
	}
}
OSF.DDA.SafeArray.Delegate.executeAsync=function OSF_DDA_SafeArray_Delegate$ExecuteAsync(args) {
	try {
		if (args.onCalling) {
			args.onCalling();
		}
		function toArray(args) {
			var arrArgs=args;
			if (OSF.OUtil.isArray(args)) {
				var len=arrArgs.length;
				for (var i=0; i < len; i++) {
					arrArgs[i]=toArray(arrArgs[i]);
				}
			} else if (OSF.OUtil.isDate(args)) {
				arrArgs=args.getVarDate();
			} else if (typeof args==="object" && !OSF.OUtil.isArray(args)) {
				arrArgs=[];
				for (var index in args) {
					if (!OSF.OUtil.isFunction(args[index])) {
						arrArgs[index]=toArray(args[index]);
					}
				}
			}
			return arrArgs;
		}
		window.external.Execute(
			args.dispId,
			toArray(args.hostCallArgs),
			function OSF_DDA_SafeArrayFacade$Execute_OnResponse(hostResponseArgs) {
				if (args.onReceiving) {
					args.onReceiving();
				}
				if (args.onComplete) {
					var result=hostResponseArgs.toArray();
					var status=result[OSF.DDA.SafeArray.Response.Status];
					var payload;
					if (status==OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess) {
						if (result.length > 2) {
							payload=[];
							for (var i=1; i < result.length; i++)
								payload[i - 1]=result[i];
						}
						else {
							payload=result[OSF.DDA.SafeArray.Response.Payload];
						}
					}
					else {
						payload=result[OSF.DDA.SafeArray.Response.Payload];
					}
					args.onComplete(status, payload);
				}
			}
		);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent=function OSF_DDA_SafeArrayDelegate$GetOnAfterRegisterEvent(args) {
	return function OSF_DDA_SafeArrayDelegate$OnAfterRegisterEvent(hostResponseArgs) {
		if (args.onReceiving) {
			args.onReceiving();
		}
		if (args.onComplete) {
			var status=hostResponseArgs.toArray ? hostResponseArgs.toArray()[OSF.DDA.SafeArray.Response.Status] : hostResponseArgs;
			args.onComplete(status)
		}
	}
}
OSF.DDA.SafeArray.Delegate.registerEventAsync=function OSF_DDA_SafeArray_Delegate$RegisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(args);
	try {
		window.external.RegisterEvent(
			args.dispId,
			args.targetId,
			function OSF_DDA_SafeArrayDelegate$RegisterEventAsync_OnEvent(eventDispId, payload) {
				if (args.onEvent) {
					args.onEvent(payload);
				}
			},
			callback
		);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
OSF.DDA.SafeArray.Delegate.unregisterEventAsync=function OSF_DDA_SafeArray_Delegate$UnregisterEventAsync(args) {
	if (args.onCalling) {
		args.onCalling();
	}
	var callback=OSF.DDA.SafeArray.Delegate._getOnAfterRegisterEvent(args);
	try {
		window.external.UnregisterEvent(
			args.dispId,
			args.targetId,
			callback
		);
	}
	catch (ex) {
		OSF.DDA.SafeArray.Delegate._onException(ex, args);
	}
};
Microsoft.Office.WebExtension.ProjectTaskFields={
	ActualCost: 0,
	ActualDuration: 1,
	ActualFinish: 2,
	ActualOvertimeCost: 3,
	ActualOvertimeWork: 4,
	ActualStart: 5,
	ActualWork: 6,
	Text1: 7,
	Text10: 8,
	Finish10: 9,
	Start10: 10,
	Text11: 11,
	Text12: 12,
	Text13: 13,
	Text14: 14,
	Text15: 15,
	Text16: 16,
	Text17: 17,
	Text18: 18,
	Text19: 19,
	Finish1: 20,
	Start1: 21,
	Text2: 22,
	Text20: 23,
	Text21: 24,
	Text22: 25,
	Text23: 26,
	Text24: 27,
	Text25: 28,
	Text26: 29,
	Text27: 30,
	Text28: 31,
	Text29: 32,
	Finish2: 33,
	Start2: 34,
	Text3: 35,
	Text30: 36,
	Finish3: 37,
	Start3: 38,
	Text4: 39,
	Finish4: 40,
	Start4: 41,
	Text5: 42,
	Finish5: 43,
	Start5: 44,
	Text6: 45,
	Finish6: 46,
	Start6: 47,
	Text7: 48,
	Finish7: 49,
	Start7: 50,
	Text8: 51,
	Finish8: 52,
	Start8: 53,
	Text9: 54,
	Finish9: 55,
	Start9: 56,
	Baseline10BudgetCost: 57,
	Baseline10BudgetWork: 58,
	Baseline10Cost: 59,
	Baseline10Duration: 60,
	Baseline10Finish: 61,
	Baseline10FixedCost: 62,
	Baseline10FixedCostAccrual: 63,
	Baseline10Start: 64,
	Baseline10Work: 65,
	Baseline1BudgetCost: 66,
	Baseline1BudgetWork: 67,
	Baseline1Cost: 68,
	Baseline1Duration: 69,
	Baseline1Finish: 70,
	Baseline1FixedCost: 71,
	Baseline1FixedCostAccrual: 72,
	Baseline1Start: 73,
	Baseline1Work: 74,
	Baseline2BudgetCost: 75,
	Baseline2BudgetWork: 76,
	Baseline2Cost: 77,
	Baseline2Duration: 78,
	Baseline2Finish: 79,
	Baseline2FixedCost: 80,
	Baseline2FixedCostAccrual: 81,
	Baseline2Start: 82,
	Baseline2Work: 83,
	Baseline3BudgetCost: 84,
	Baseline3BudgetWork: 85,
	Baseline3Cost: 86,
	Baseline3Duration: 87,
	Baseline3Finish: 88,
	Baseline3FixedCost: 89,
	Baseline3FixedCostAccrual: 90,
	Basline3Start: 91,
	Baseline3Work: 92,
	Baseline4BudgetCost: 93,
	Baseline4BudgetWork: 94,
	Baseline4Cost: 95,
	Baseline4Duration: 96,
	Baseline4Finish: 97,
	Baseline4FixedCost: 98,
	Baseline4FixedCostAccrual: 99,
	Baseline4Start: 100,
	Baseline4Work: 101,
	Baseline5BudgetCost: 102,
	Baseline5BudgetWork: 103,
	Baseline5Cost: 104,
	Baseline5Duration: 105,
	Baseline5Finish: 106,
	Baseline5FixedCost: 107,
	Baseline5FixedCostAccrual: 108,
	Baseline5Start: 109,
	Baseline5Work: 110,
	Baseline6BudgetCost: 111,
	Baseline6BudgetWork: 112,
	Baseline6Cost: 113,
	Baseline6Duration: 114,
	Baseline6Finish: 115,
	Baseline6FixedCost: 116,
	Baseline6FixedCostAccrual: 117,
	Baseline6Start: 118,
	Baseline6Work: 119,
	Baseline7BudgetCost: 120,
	Baseline7BudgetWork: 121,
	Baseline7Cost: 122,
	Baseline7Duration: 123,
	Baseline7Finish: 124,
	Baseline7FixedCost: 125,
	Baseline7FixedCostAccrual: 126,
	Baseline7Start: 127,
	Baseline7Work: 128,
	Baseline8BudgetCost: 129,
	Baseline8BudgetWork: 130,
	Baseline8Cost: 131,
	Baseline8Duration: 132,
	Baseline8Finish: 133,
	Baseline8FixedCost: 134,
	Baseline8FixedCostAccrual: 135,
	Baseline8Start: 136,
	Baseline8Work: 137,
	Baseline9BudgetCost: 138,
	Baseline9BudgetWork: 139,
	Baseline9Cost: 140,
	Baseline9Duration: 141,
	Baseline9Finish: 142,
	Baseline9FixedCost: 143,
	Baseline9FixedCostAccrual: 144,
	Baseline9Start: 145,
	Baseline9Work: 146,
	BaselineBudgetCost: 147,
	BaselineBudgetWork: 148,
	BaselineCost: 149,
	BaselineDuration: 150,
	BaselineFinish: 151,
	BaselineFixedCost: 152,
	BaselineFixedCostAccrual: 153,
	BaselineStart: 154,
	BaselineWork: 155,
	BudgetCost: 156,
	BudgetFixedCost: 157,
	BudgetFixedWork: 158,
	BudgetWork: 159,
	TaskCalendarGUID: 160,
	ConstraintDate: 161,
	ConstraintType: 162,
	Cost1: 163,
	Cost10: 164,
	Cost2: 165,
	Cost3: 166,
	Cost4: 167,
	Cost5: 168,
	Cost6: 169,
	Cost7: 170,
	Cost8: 171,
	Cost9: 172,
	Date1: 173,
	Date10: 174,
	Date2: 175,
	Date3: 176,
	Date4: 177,
	Date5: 178,
	Date6: 179,
	Date7: 180,
	Date8: 181,
	Date9: 182,
	Deadline: 183,
	Duration1: 184,
	Duration10: 185,
	Duration2: 186,
	Duration3: 187,
	Duration4: 188,
	Duration5: 189,
	Duration6: 190,
	Duration7: 191,
	Duration8: 192,
	Duration9: 193,
	Duration: 194,
	EarnedValueMethod: 195,
	FinishSlack: 196,
	FixedCost: 197,
	FixedCostAccrual: 198,
	Flag10: 199,
	Flag1: 200,
	Flag11: 201,
	Flag12: 202,
	Flag13: 203,
	Flag14: 204,
	Flag15: 205,
	Flag16: 206,
	Flag17: 207,
	Flag18: 208,
	Flag19: 209,
	Flag2: 210,
	Flag20: 211,
	Flag3: 212,
	Flag4: 213,
	Flag5: 214,
	Flag6: 215,
	Flag7: 216,
	Flag8: 217,
	Flag9: 218,
	FreeSlack: 219,
	HasRollupSubTasks: 220,
	ID: 221,
	Name: 222,
	Notes: 223,
	Number1: 224,
	Number10: 225,
	Number11: 226,
	Number12: 227,
	Number13: 228,
	Number14: 229,
	Number15: 230,
	Number16: 231,
	Number17: 232,
	Number18: 233,
	Number19: 234,
	Number2: 235,
	Number20: 236,
	Number3: 237,
	Number4: 238,
	Number5: 239,
	Number6: 240,
	Number7: 241,
	Number8: 242,
	Number9: 243,
	ScheduledDuration: 244,
	ScheduledFinish: 245,
	ScheduledStart: 246,
	OutlineLevel: 247,
	OvertimeCost: 248,
	OvertimeWork: 249,
	PercentComplete: 250,
	PercentWorkComplete: 251,
	Predecessors: 252,
	PreleveledFinish: 253,
	PreleveledStart: 254,
	Priority: 255,
	Active: 256,
	Critical: 257,
	Milestone: 258,
	Overallocated: 259,
	IsRollup: 260,
	Summary: 261,
	RegularWork: 262,
	RemainingCost: 263,
	RemainingDuration: 264,
	RemainingOvertimeCost: 265,
	RemainingWork: 266,
	ResourceNames: 267,
	ResourceNames: 268,
	Cost: 269,
	Finish: 270,
	Start: 271,
	Work: 272,
	StartSlack: 273,
	Status: 274,
	Successors: 275,
	StatusManager: 276,
	TotalSlack: 277,
	TaskGUID: 278,
	Type: 279,
	WBS: 280,
	WBSPREDECESSORS: 281,
	WBSSUCCESSORS: 282,
	WSSID: 283
}
Microsoft.Office.WebExtension.ProjectResourceFields={
	Accrual: 0,
	ActualCost: 1,
	ActualOvertimeCost: 2,
	ActualOvertimeWork: 3,
	ActualOvertimeWorkProtected: 4,
	ActualWork: 5,
	ActualWorkProtected: 6,
	BaseCalendar: 7,
	Baseline10BudgetCost: 8,
	Baseline10BudgetWork: 9,
	Baseline10Cost: 10,
	Baseline10Work: 11,
	Baseline1BudgetCost: 12,
	Baseline1BudgetWork: 13,
	Baseline1Cost: 14,
	Baseline1Work: 15,
	Baseline2BudgetCost: 16,
	Baseline2BudgetWork: 17,
	Baseline2Cost: 18,
	Baseline2Work: 19,
	Baseline3BudgetCost: 20,
	Baseline3BudgetWork: 21,
	Baseline3Cost: 22,
	Baseline3Work: 23,
	Baseline4BudgetCost: 24,
	Baseline4BudgetWork: 25,
	Baseline4Cost: 26,
	Baseline4Work: 27,
	Baseline5BudgetCost: 28,
	Baseline5BudgetWork: 29,
	Baseline5Cost: 30,
	Baseline5Work: 31,
	Baseline6BudgetCost: 32,
	Baseline6BudgetWork: 33,
	Baseline6Cost: 34,
	Baseline6Work: 35,
	Baseline7BudgetCost: 36,
	Baseline7BudgetWork: 37,
	Baseline7Cost: 38,
	Baseline7Work: 39,
	Baseline8BudgetCost: 40,
	Baseline8BudgetWork: 41,
	Baseline8Cost: 42,
	Baseline8Work: 43,
	Baseline9BudgetCost: 44,
	Baseline9BudgetWork: 45,
	Baseline9Cost: 46,
	Baseline9Work: 47,
	BaselineBudgetCost: 48,
	BaselineBudgetWork: 49,
	BaselineCost: 50,
	BaselineWork: 51,
	BudgetCost: 52,
	BudgetWork: 53,
	ResourceCalendarGUID: 54,
	Code: 55,
	Cost1: 56,
	Cost10: 57,
	Cost2: 58,
	Cost3: 59,
	Cost4: 60,
	Cost5: 61,
	Cost6: 62,
	Cost7: 63,
	Cost8: 64,
	Cost9: 65,
	ResourceCreationDate: 66,
	Date1: 67,
	Date10: 68,
	Date2: 69,
	Date3: 70,
	Date4: 71,
	Date5: 72,
	Date6: 73,
	Date7: 74,
	Date8: 75,
	Date9: 76,
	Duration1: 77,
	Duration10: 78,
	Duration2: 79,
	Duration3: 80,
	Duration4: 81,
	Duration5: 82,
	Duration6: 83,
	Duration7: 84,
	Duration8: 85,
	Duration9: 86,
	Email: 87,
	End: 88,
	Finish1: 89,
	Finish10: 90,
	Finish2: 91,
	Finish3: 92,
	Finish4: 93,
	Finish5: 94,
	Finish6: 95,
	Finish7: 96,
	Finish8: 97,
	Finish9: 98,
	Flag10: 99,
	Flag1: 100,
	Flag11: 101,
	Flag12: 102,
	Flag13: 103,
	Flag14: 104,
	Flag15: 105,
	Flag16: 106,
	Flag17: 107,
	Flag18: 108,
	Flag19: 109,
	Flag2: 110,
	Flag20: 111,
	Flag3: 112,
	Flag4: 113,
	Flag5: 114,
	Flag6: 115,
	Flag7: 116,
	Flag8: 117,
	Flag9: 118,
	Group: 119,
	Units: 120,
	Name: 121,
	Notes: 122,
	Number1: 123,
	Number10: 124,
	Number11: 125,
	Number12: 126,
	Number13: 127,
	Number14: 128,
	Number15: 129,
	Number16: 130,
	Number17: 131,
	Number18: 132,
	Number19: 133,
	Number2: 134,
	Number20: 135,
	Number3: 136,
	Number4: 137,
	Number5: 138,
	Number6: 139,
	Number7: 140,
	Number8: 141,
	Number9: 142,
	OvertimeCost: 143,
	OvertimeRate: 144,
	OvertimeWork: 145,
	PercentWorkComplete: 146,
	CostPerUse: 147,
	Generic: 148,
	OverAllocated: 149,
	RegularWork: 150,
	RemainingCost: 151,
	RemainingOvertimeCost: 152,
	RemainingOvertimeWork: 153,
	RemainingWork: 154,
	ResourceGUID: 155,
	Cost: 156,
	Work: 157,
	Start: 158,
	Start1: 159,
	Start10: 160,
	Start2: 161,
	Start3: 162,
	Start4: 163,
	Start5: 164,
	Start6: 165,
	Start7: 166,
	Start8: 167,
	Start9: 168,
	StandardRate: 169,
	Text1: 170,
	Text10: 171,
	Text11: 172,
	Text12: 173,
	Text13: 174,
	Text14: 175,
	Text15: 176,
	Text16: 177,
	Text17: 178,
	Text18: 179,
	Text19: 180,
	Text2: 181,
	Text20: 182,
	Text21: 183,
	Text22: 184,
	Text23: 185,
	Text24: 186,
	Text25: 187,
	Text26: 188,
	Text27: 189,
	Text28: 190,
	Text29: 191,
	Text3: 192,
	Text30: 193,
	Text4: 194,
	Text5: 195,
	Text6: 196,
	Text7: 197,
	Text8: 198,
	Text9: 199
}
Microsoft.Office.WebExtension.ProjectProjectFields={
	CurrencyDigits: 0,
	CurrencySymbol: 1,
	CurrencySymbolPosition: 2,
	DurationUnits: 3,
	GUID: 4,
	Finish: 5,
	Start: 6,
	ReadOnly: 7,
	VERSION: 8,
	WorkUnits: 9,
	ProjectServerUrl: 10,
	WSSUrl: 11,
	WSSList: 12
}
Microsoft.Office.WebExtension.ProjectViewTypes={
	Gantt           : 1,
	NetworkDiagram  : 2,
	TaskDiagram     : 3,
	TaskForm        : 4,
	TaskSheet       : 5,
	ResourceForm    : 6,
	ResourceSheet   : 7,
	ResourceGraph   : 8,
	TeamPlanner     : 9,
	TaskDetails     : 10,
	TaskNameForm    : 11,
	ResourceNames   : 12,
	Calendar        : 13,
	TaskUsage       : 14,
	ResourceUsage   : 15,
	Timeline        : 16
}
Microsoft.Office.WebExtension.CoercionType={
	Text: "text"
}
Microsoft.Office.WebExtension.ValueFormat={
	Unformatted: "unformatted"
}
Microsoft.Office.WebExtension.FilterType={
	All: "all"
}
Microsoft.Office.WebExtension.EventType.TaskSelectionChanged="taskSelectionChanged";
Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged="resourceSelectionChanged";
Microsoft.Office.WebExtension.EventType.ViewSelectionChanged="viewSelectionChanged";
OSF.DDA.ProjectDocument=function OSF_DDA_ProjectDocument(officeAppContext) {
	OSF.DDA.ProjectDocument.uber.constructor.call(this,
		officeAppContext
	);
	var am=OSF.DDA.AsyncMethodNames;
	OSF.DDA.DispIdHost.addAsyncMethods(this, [
		am.GetSelectedDataAsync,
		am.GetSelectedTask,
		am.GetTask,
		am.GetTaskField,
		am.GetWSSUrl,
		am.GetSelectedResource,
		am.GetResourceField,
		am.GetProjectField,
		am.GetSelectedView
	]);
	OSF.DDA.DispIdHost.addEventSupport(this, new OSF.EventDispatch([Microsoft.Office.WebExtension.EventType.TaskSelectionChanged,
																	Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged,
																	Microsoft.Office.WebExtension.EventType.ViewSelectionChanged]));
}
OSF.OUtil.extend(OSF.DDA.ProjectDocument, OSF.DDA.Document);
OSF.DDA.TaskSelectionChangedEventArgs=function OSF_DDA_TaskSelectionChangedEventArgs(doc) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.TaskSelectionChanged
		},
		"document": {
			value: doc
		}
	});
}
OSF.DDA.ResourceSelectionChangedEventArgs=function OSF_DDA_ResourceSelectionChangedEventArgs(doc) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.ResourceSelectionChanged
		},
		"document": {
			value: doc
		}
	});
}
OSF.DDA.ViewSelectionChangedEventArgs=function OSF_DDA_ViewSelectionChangedEventArgs(doc) {
	OSF.OUtil.defineEnumerableProperties(this, {
		"type": {
			value: Microsoft.Office.WebExtension.EventType.ViewSelectionChanged
		},
		"document": {
			value: doc
		}
	});
}

