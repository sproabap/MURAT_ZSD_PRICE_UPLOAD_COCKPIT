CLASS zcl_sd_fields_validations DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    TYPES: BEGIN OF ts_check_fields,
             index             TYPE sy-tabix,
             salesorganization TYPE vkorg,
             customer          TYPE kunag,
             operationtype     TYPE vtweg,
             departureprovince TYPE zsd_departureprovince,
             departuredistrict TYPE zsd_departuredistrict,
             originpoint       TYPE zsd_origin_point,
             arrivalprovince   TYPE zsd_arrivalprovince,
             arrivaldistrict   TYPE zsd_arrivaldistrict,
             destinationpoint  TYPE zsd_destination_point,
             routetype         TYPE zsd_routetype,
             routecode         TYPE zsd_route_code,
             vehicletype       TYPE zsd_vehicletype,
             vehiclebodytype   TYPE zsd_vehiclebodytype,
             temperatureregime TYPE zsd_temperatureregime,
             packettype        TYPE zsd_packagetype,
             conditionunit     TYPE meins,
             teamdriverstatus  TYPE zsd_teamdriverstatus,
           END OF ts_check_fields.
    TYPES tt_check_fields TYPE STANDARD TABLE OF ts_check_fields WITH DEFAULT KEY.

    CLASS-DATA gt_check_fields   TYPE tt_check_fields.
    CLASS-DATA gt_check_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS check_fields
      IMPORTING iv_table_type     TYPE zr_sd_t_excel-Tabletype
      EXPORTING et_check_response TYPE zcl_sd_types=>tt_response
      CHANGING  ct_check_fields   TYPE ANY TABLE.

  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_SD_FIELDS_VALIDATIONS IMPLEMENTATION.


  METHOD check_fields.
    CLEAR gt_check_fields.
    CLEAR gt_check_response.
    MOVE-CORRESPONDING ct_check_fields TO gt_check_fields.
    DATA lv_index TYPE sy-tabix.

    LOOP AT gt_check_fields ASSIGNING FIELD-SYMBOL(<fs_check_fields>).
      lv_index += 1.
      <fs_check_fields>-index = lv_index.
*      IF <fs_check_fields>-teamdriverstatus <> 'X'.
*        <fs_check_fields>-teamdriverstatus = ''.
*      ENDIF.
    ENDLOOP.


    SELECT *
      FROM @gt_check_fields AS fields
             LEFT JOIN
               I_Customer AS customer ON customer~customer = fields~customer
                 LEFT JOIN
                   I_DistributionChannel AS distribution ON distribution~DistributionChannel = fields~operationtype
                     LEFT JOIN
                       zi_departureprovinces AS departureprovince ON departureprovince~departureprovince = fields~departureprovince
                         LEFT JOIN
                           zi_departuredistricts AS departuredistricts ON departuredistricts~DepartureDistrict = fields~departuredistrict
                             LEFT JOIN
                               zi_originpoint AS origin ON origin~OriginPoint = fields~originpoint
                                 LEFT JOIN
                                   zi_arrivalprovinces AS arrivalprovinces ON arrivalprovinces~ArrivalProvince = fields~arrivalprovince
                                     LEFT JOIN
                                       zi_departuredistricts AS arrivaldistrict ON arrivaldistrict~DepartureDistrict = fields~arrivaldistrict
                                         LEFT JOIN
                                           zi_destinationpoint AS destination ON destination~DestinationPoint = fields~destinationpoint
                                             LEFT JOIN
                                               zi_routetypes AS routetype ON routetype~RouteType = fields~routetype
                                                 LEFT JOIN
                                                   zi_vehicletype AS vehicle ON vehicle~VehicleType = fields~vehicletype
                                                     LEFT JOIN
                                                       zi_vehiclebodytype AS vehiclebody ON vehiclebody~VehicleBodyType = fields~vehiclebodytype
                                                         LEFT JOIN
                                                           zi_temperatureregime AS temp ON temp~temperatureregime = fields~temperatureregime
                                                             LEFT JOIN
                                                               zi_route AS routecode ON routecode~route = fields~routecode
                                                                 LEFT JOIN
                                                                   zpackagetype AS package ON package~PackageType = fields~packettype
                                                                     LEFT JOIN
                                                                       I_UnitOfMeasure AS units ON units~UnitOfMeasure = fields~conditionunit
                                                                         LEFT JOIN
                                                                           I_SalesOrganization AS sales ON sales~SalesOrganization = fields~salesorganization
      ORDER BY index
      INTO TABLE @DATA(lt_check).

    SELECT *                                  "#EC CI_ALL_FIELDS_NEEDED
      FROM ddcds_customer_domain_value_t( p_domain_name = 'ZSD_TABLE_FIELDS' )
      WHERE language = @sy-langu
      INTO TABLE @DATA(lt_fields).

    CLEAR lv_index.
    LOOP AT lt_check INTO DATA(ls_check).
      lv_index += 1.
      IF ls_check-units IS INITIAL AND iv_table_type <> '03'.
        APPEND VALUE #( flag  = '8'
                        tabix = lv_index + 1 ) TO gt_check_response.
      ENDIF.
      IF ls_check-customer IS INITIAL.
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '01' ]-text ) TO gt_check_response.
      ENDIF.
      IF     ls_check-distribution IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '05' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '02' ]-text  ) TO gt_check_response.
      ENDIF.
      IF     ls_check-departuredistricts IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '05' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '03' ]-text  ) TO gt_check_response.
      ENDIF.
      IF     ls_check-departureprovince IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '05' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '04' ]-text  ) TO gt_check_response.
      ENDIF.
      IF     ls_check-origin IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '05' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '05' ]-text  ) TO gt_check_response.
      ENDIF.
      IF     ls_check-routetype IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '05' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '06' ]-text  ) TO gt_check_response.
      ENDIF.
      IF ls_check-vehicle IS INITIAL.
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '07' ]-text  ) TO gt_check_response.
      ENDIF.
      IF ls_check-vehiclebody IS INITIAL.
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '08' ]-text  ) TO gt_check_response.
      ENDIF.
      IF ls_check-temp IS INITIAL.
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '09' ]-text  ) TO gt_check_response.
      ENDIF.
      IF     ls_check-arrivaldistrict IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '10' ]-text  ) TO gt_check_response.
      ENDIF.
      IF     ls_check-arrivalprovinces IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '11' ]-text ) TO gt_check_response.
      ENDIF.
      IF     ls_check-destination IS INITIAL
         AND ( iv_table_type = '01' OR iv_table_type = '02' OR iv_table_type = '03' OR iv_table_type = '04' OR iv_table_type = '08' OR iv_table_type = '09' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_low = '12' ]-text ) TO gt_check_response.
      ENDIF.
      IF ls_check-routecode IS INITIAL AND ( iv_table_type = '05' ).
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_position = '13' ]-text  ) TO gt_check_response.
      ENDIF.
      IF ls_check-package IS INITIAL AND iv_table_type = '03'.
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_position = '14' ]-text  ) TO gt_check_response.

      ENDIF.
      IF ls_check-sales IS INITIAL.
        APPEND VALUE #( flag  = '3'
                        tabix = lv_index + 1
                        field = lt_fields[ value_position = '15' ]-text  ) TO gt_check_response.
      ENDIF.
    ENDLOOP.

    MOVE-CORRESPONDING gt_check_response TO et_check_response.
*    MOVE-CORRESPONDING gt_check_fields TO ct_check_fields.
  ENDMETHOD.
ENDCLASS.
