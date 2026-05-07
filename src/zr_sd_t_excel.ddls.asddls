@AccessControl.authorizationCheck: #CHECK
@Metadata.allowExtensions: true
@EndUserText.label: 'Price Upload Cockpit'
define root view entity ZR_SD_T_EXCEL
  as select from zsd_t_excel
  association to ZSD_I_TABLETYPES_VH    as _Table      on  _Table.value_low = $projection.Tabletype
                                                       and _Table.language  = $session.system_language
  association to I_BusinessUserBasic    as _UserCreate on  _UserCreate.UserID = $projection.CreatedBy
  association to I_BusinessUserBasic    as _UserChange on  _UserChange.UserID = $projection.LastChangedBy
  association to ZI_SdTableTypesMainten as _Maint      on  _Maint.Tabletype = $projection.Tabletype
{
  key attachid              as Attachid,
  key tabletype             as Tabletype,
      @Semantics.largeObject: { mimeType: 'Mimetype',
                              fileName: 'Filename',
                              contentDispositionPreference: #ATTACHMENT,
                              acceptableMimeTypes: [ 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ]
                              }
      attachment            as Attachment,
      @Semantics.mimeType: true
      mimetype              as Mimetype,
      filename              as Filename,
      @Semantics.user.createdBy: true
      created_by            as CreatedBy,
      @Semantics.systemDateTime.createdAt: true
      created_at            as CreatedAt,
      @Semantics.user.lastChangedBy: true
      last_changed_by       as LastChangedBy,
      @Semantics.systemDateTime.lastChangedAt: true
      last_changed_at       as LastChangedAt,
      @Semantics.systemDateTime.localInstanceLastChangedAt: true
      local_last_changed_at as LocalLastChangedAt,

      _Table,
      _UserCreate,
      _UserChange,



            '#BusinessConfiguration-maintain&/Detail%20(ZSDTABLETYPESMAINTEN)&/ZSDTABLETYPESMAINTEN/SdTableTypesMainAll(1)' as url
//      concat(
//              concat(
//                  'https://my418838.s4hana.cloud.sap/ui#BusinessConfiguration-maintain&/Detail%20(ZSDTABLETYPESMAINTEN)&/ZSDTABLETYPESMAINTEN/SdTableTypesMainAll(SingletonID=1,IsActiveEntity=true)/_SdTableTypesMainten(Attachid=',
//                  '_Maint.Attachid'
//              ),
//              concat(
//                  ',Tabletype=',
//                  $projection.Tabletype
//              )
//          )                 as url
      //    concat(concat(concat('#BusinessConfiguration-maintain&/Detail%20(ZSDTABLETYPESMAINTEN)&/ZSDTABLETYPESMAINTEN/SdTableTypesMainAll(SingletonID=1,IsActiveEntity=true)/_SdTableTypesMainten(Attachid=',_Maint.Attachid),'Tabletype='''), $projection.Tabletype),'''IsActiveEntity=true))') as url
}
