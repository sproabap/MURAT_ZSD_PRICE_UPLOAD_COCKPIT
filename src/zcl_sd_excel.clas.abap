CLASS zcl_sd_excel DEFINITION
  PUBLIC FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    CLASS-DATA gt_shipping_based TYPE zcl_sd_types=>tt_shipping_based.
    CLASS-DATA gt_km_based       TYPE zcl_sd_types=>tt_km_based.
    CLASS-DATA gt_pallet_based   TYPE zcl_sd_types=>tt_pallet_based.
    CLASS-DATA gt_weight_based   TYPE zcl_sd_types=>tt_weight_based.
    CLASS-DATA gt_route_based    TYPE zcl_sd_types=>tt_route_based.
    CLASS-DATA gt_fixedrent      TYPE zcl_sd_types=>tt_fixedrent.
    CLASS-DATA gt_pointdiff      TYPE zcl_sd_types=>tt_pointdiff.
    CLASS-DATA gt_kmdiff         TYPE zcl_sd_types=>tt_kmdiff.
    CLASS-DATA gt_porterage      TYPE zcl_sd_types=>tt_porterage.
    CLASS-DATA gt_waiting        TYPE zcl_sd_types=>tt_waiting.
    CLASS-DATA gt_tollcost       TYPE zcl_sd_types=>tt_tollcost.

    CLASS-METHODS read_shipping_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_km_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_pallet_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_weight_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_route_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_fixedrent_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_pointdiff_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_kmdiff_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_porterage_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_waiting_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.

    CLASS-METHODS read_tollcost_excel
      IMPORTING attachment  TYPE xstring
      EXPORTING et_response TYPE zcl_sd_types=>tt_response.
ENDCLASS.



CLASS ZCL_SD_EXCEL IMPLEMENTATION.


  METHOD read_fixedrent_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'N' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_fixedrent[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_fixedrent WHERE customer = ''.

    LOOP AT gt_fixedrent ASSIGNING FIELD-SYMBOL(<fs_fixedrent>).
      DATA(lv_tabix) = sy-tabix.
      IF <fs_fixedrent>-teamdriversts <> 'X'.
        <fs_fixedrent>-teamdriversts = ''.
      ENDIF.
      LOOP AT gt_fixedrent TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer        = <fs_fixedrent>-customer        AND vehicletype       = <fs_fixedrent>-vehicletype
                                                                          AND vehiclebodytype = <fs_fixedrent>-vehiclebodytype AND temperatureregime = <fs_fixedrent>-temperatureregime
                                                                          AND teamdriversts   = <fs_fixedrent>-teamdriversts   AND operationtype     = <fs_fixedrent>-operationtype
                                                                          AND salesorganization = <fs_fixedrent>-salesorganization
                                                                          AND ( (     validfrom = <fs_fixedrent>-validfrom
                                                                                  AND validto   = <fs_fixedrent>-validto ) OR ( validto > <fs_fixedrent>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_fixedrent>-customer      = |{ <fs_fixedrent>-customer ALPHA = IN }|.
      <fs_fixedrent>-conditionunit = 'MON'.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '06'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_fixedrent ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_fixedrentmdl.

    IF gt_fixedrent IS NOT INITIAL.

      SELECT fixedrent~client,
             fixedrent~uuid,
             fixedrent~salesorganization,
             fixedrent~customer,
             fixedrent~teamdriversts,
             fixedrent~vehicletype,
             fixedrent~vehiclebodytype,
             fixedrent~temperatureregime,
             fixedrent~daysworked,
             fixedrent~conditionquantity,
             fixedrent~conditionunit,
             fixedrent~conditionprice,
             fixedrent~currency,
             fixedrent~validfrom,
             fixedrent~validto
        FROM zsd_fixedrentmdl AS fixedrent
             INNER JOIN
             @gt_fixedrent    AS fixed     ON  fixedrent~customer          = fixed~Customer          AND fixedrent~teamdriversts   = fixed~teamdriversts
                                           AND fixedrent~vehicletype       = fixed~Vehicletype       AND fixedrent~vehiclebodytype = fixed~Vehiclebodytype
                                           AND fixedrent~temperatureregime = fixed~Temperatureregime AND fixedrent~salesorganization = fixed~salesorganization
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_fixedrentmdl.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_fixedrentmdl.

    MOVE-CORRESPONDING gt_fixedrent TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control = VALUE #( salesorganization = '01'
                                      customer          = '01'
                                      Operationtype     = '01'
                                      Teamdriversts     = '01'
                                      vehicletype       = '01'
                                      vehiclebodytype   = '01'
                                      Temperatureregime = '01'
                                      daysworked        = '01'
                                      Conditionquantity = '01'
                                      Conditionunit     = '01'
                                      Conditionprice    = '01'
                                      Currency          = '01'
                                      validfrom         = '01'
                                      Validto           = '01' ).
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_fixedrentmdl
             ENTITY FixedRentModel CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ELSE.
      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        DATA(lv_date_old_start) = gt_fixedrent[ salesorganization = <fs_update>-Salesorganization
                                                customer          = <fs_update>-Customer
                                                teamdriversts     = <fs_update>-Teamdriversts
                                                vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                vehicletype       = <fs_update>-Vehicletype
                                                temperatureregime = <fs_update>-Temperatureregime ]-ValidFrom .

        DATA(lv_date_old_end) = gt_fixedrent[ salesorganization = <fs_update>-Salesorganization
                                              customer          = <fs_update>-Customer
                                              teamdriversts     = <fs_update>-Teamdriversts
                                              vehiclebodytype   = <fs_update>-Vehiclebodytype
                                              vehicletype       = <fs_update>-Vehicletype
                                              temperatureregime = <fs_update>-Temperatureregime ]-validto.

        IF lv_date_old_start <= <fs_update>-Validto AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validto = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validto AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_fixedrentmdl
               ENTITY FixedRentModel UPDATE FIELDS ( Validfrom Validto ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_fixedrentmdl
               ENTITY FixedRentModel CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_kmdiff_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'S' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_kmdiff[] )
      )->if_xco_xlsx_ra_operation~execute( ).

    LOOP AT gt_kmdiff ASSIGNING FIELD-SYMBOL(<fs_kmdiff>).
      DATA(lv_tabix) = sy-tabix.
      LOOP AT gt_kmdiff TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_kmdiff>-customer          AND operationtype     = <fs_kmdiff>-operationtype
                                                                       AND departureprovince = <fs_kmdiff>-departureprovince AND departuredistrict = <fs_kmdiff>-departuredistrict
                                                                       AND originpoint       = <fs_kmdiff>-originpoint       AND arrivaldistrict   = <fs_kmdiff>-arrivaldistrict
                                                                       AND arrivalprovince   = <fs_kmdiff>-arrivalprovince   AND destinationpoint  = <fs_kmdiff>-destinationpoint
                                                                       AND routetype         = <fs_kmdiff>-routetype         AND vehicletype       = <fs_kmdiff>-vehicletype
                                                                       AND vehiclebodytype   = <fs_kmdiff>-vehiclebodytype   AND temperatureregime = <fs_kmdiff>-temperatureregime
                                                                       AND salesorganization = <fs_kmdiff>-salesorganization
                                                                       AND ( (     validfrom   = <fs_kmdiff>-validfrom
                                                                               AND validtodate = <fs_kmdiff>-validtodate ) OR ( validtodate > <fs_kmdiff>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_kmdiff>-customer = |{ <fs_kmdiff>-customer ALPHA = IN }|.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '08'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_kmdiff ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_t_kmdiff.

    IF gt_kmdiff IS NOT INITIAL.

      SELECT kmdiff~client,
             kmdiff~uuid,
             kmdiff~salesorganization,
             kmdiff~customer,
             kmdiff~operationtype,
             kmdiff~departureprovince,
             kmdiff~departuredistrict,
             kmdiff~originpoint,
             kmdiff~arrivalprovince,
             kmdiff~arrivaldistrict,
             kmdiff~destinationpoint,
             kmdiff~routetype,
             kmdiff~vehicletype,
             kmdiff~vehiclebodytype,
             kmdiff~temperatureregime,
             kmdiff~conditionquantity,
             kmdiff~conditionunit,
             kmdiff~conditionprice,
             kmdiff~currency,
             kmdiff~validfrom,
             kmdiff~validtodate
        FROM zsd_t_kmdiff AS kmdiff
             INNER JOIN
             @gt_kmdiff   AS km     ON  kmdiff~customer          = km~Customer          AND kmdiff~operationtype     = km~Operationtype
                                    AND kmdiff~arrivalprovince   = km~Arrivalprovince   AND kmdiff~arrivaldistrict   = km~Arrivaldistrict
                                    AND kmdiff~originpoint       = km~Originpoint       AND kmdiff~departuredistrict = km~Departuredistrict
                                    AND kmdiff~destinationpoint  = km~Destinationpoint  AND kmdiff~routetype         = km~Routetype
                                    AND kmdiff~vehicletype       = km~Vehicletype       AND kmdiff~vehiclebodytype   = km~Vehiclebodytype
                                    AND kmdiff~temperatureregime = km~Temperatureregime AND kmdiff~salesorganization = km~salesorganization
        INTO TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_kmdiff.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_kmdiff.

    MOVE-CORRESPONDING gt_kmdiff TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control = VALUE #( Salesorganization = '01'
                                      Customer          = '01'
                                      Operationtype     = '01'
                                      Departureprovince = '01'
                                      Departuredistrict = '01'
                                      Originpoint       = '01'
                                      Arrivalprovince   = '01'
                                      Arrivaldistrict   = '01'
                                      Destinationpoint  = '01'
                                      Routetype         = '01'
                                      Vehicletype       = '01'
                                      Vehiclebodytype   = '01'
                                      Temperatureregime = '01'
                                      Conditionprice    = '01'
                                      Conditionquantity = '01'
                                      Conditionunit     = '01'
                                      Currency          = '01'
                                      Validfrom         = '01'
                                      Validtodate       = '01' ).
      <fs_create>-Customer = |{ <fs_create>-Customer ALPHA = IN } |.
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_t_kmdiff
             ENTITY KMDifference CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ELSE.
      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        READ TABLE gt_kmdiff INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                          customer          = <fs_update>-Customer
                                                          operationtype     = <fs_update>-Operationtype
                                                          departuredistrict = <fs_update>-Departuredistrict
                                                          departureprovince = <fs_update>-Departureprovince
                                                          originpoint       = <fs_update>-Originpoint
                                                          arrivaldistrict   = <fs_update>-Arrivaldistrict
                                                          arrivalprovince   = <fs_update>-Arrivalprovince
                                                          destinationpoint  = <fs_update>-Destinationpoint
                                                          routetype         = <fs_update>-Routetype
                                                          vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                          vehicletype       = <fs_update>-Vehicletype
                                                          temperatureregime = <fs_update>-Temperatureregime.
        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-ValidToDate.
        ENDIF.
        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.

        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_kmdiff
               ENTITY KMDifference UPDATE FIELDS ( Validfrom Validtodate ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_kmdiff
               ENTITY KMDifference CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_km_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'V' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_km_based[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_km_based WHERE customer = ''.

    LOOP AT gt_km_based ASSIGNING FIELD-SYMBOL(<fs_km_based>).
      DATA(lv_tabix) = sy-tabix.
            IF <fs_km_based>-teamdriverstatus <> 'X'.
        <fs_km_based>-teamdriverstatus = ''.
      ENDIF.
      LOOP AT gt_km_based TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_km_based>-customer          AND operationtype     = <fs_km_based>-operationtype
                                                                         AND departureprovince = <fs_km_based>-departureprovince AND departuredistrict = <fs_km_based>-departuredistrict
                                                                         AND originpoint       = <fs_km_based>-originpoint       AND arrivaldistrict   = <fs_km_based>-arrivaldistrict
                                                                         AND arrivalprovince   = <fs_km_based>-arrivalprovince   AND destinationpoint  = <fs_km_based>-destinationpoint
                                                                         AND routetype         = <fs_km_based>-routetype         AND vehicletype       = <fs_km_based>-vehicletype
                                                                         AND vehiclebodytype   = <fs_km_based>-vehiclebodytype   AND temperatureregime = <fs_km_based>-temperatureregime
                                                                         AND teamdriverstatus  = <fs_km_based>-teamdriverstatus  AND salesorganization = <fs_km_based>-salesorganization
                                                                         AND ( (     validfrom   = <fs_km_based>-validfrom
                                                                                 AND validtodate = <fs_km_based>-validtodate ) OR ( validtodate > <fs_km_based>-validfrom ) )
                                                                         AND ( kmend = <fs_km_based>-kmstart OR kmstart = <fs_km_based>-kmstart ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0 AND <fs_km_based>-kmstart <> <fs_km_based>-kmend.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_km_based>-customer = |{ <fs_km_based>-customer ALPHA = IN }|.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '02'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_km_based ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_t_kmbase.

    IF gt_km_based IS NOT INITIAL.

      SELECT DISTINCT kmbased~uuid,
                      kmbased~customer,
                      kmbased~salesorganization,
                      kmbased~operationtype,
                      kmbased~departureprovince,
                      kmbased~departuredistrict,
                      kmbased~originpoint,
                      kmbased~arrivalprovince,
                      kmbased~arrivaldistrict,
                      kmbased~destinationpoint,
                      kmbased~routetype,
                      kmbased~vehicletype,
                      kmbased~vehiclebodytype,
                      kmbased~temperatureregime,
                      kmbased~teamdriverstatus,
                      kmbased~validfrom,
                      kmbased~validtodate,
                      kmbased~kmstart,
                      kmbased~kmend
        FROM zsd_t_kmbase AS kmbased
             INNER JOIN
             @gt_km_based AS km      ON  kmbased~customer          = km~Customer          AND kmbased~salesorganization = km~salesorganization
                                     AND kmbased~operationtype     = km~Operationtype     AND kmbased~departureprovince = km~Departureprovince
                                     AND kmbased~departuredistrict = km~Departuredistrict AND kmbased~originpoint       = km~Originpoint
                                     AND kmbased~arrivalprovince   = km~Arrivalprovince   AND kmbased~arrivaldistrict   = km~Arrivaldistrict
                                     AND kmbased~destinationpoint  = km~Destinationpoint  AND kmbased~routetype         = km~Routetype
                                     AND kmbased~vehicletype       = km~Vehicletype       AND kmbased~vehiclebodytype   = km~Vehiclebodytype
                                     AND kmbased~temperatureregime = km~Temperatureregime AND kmbased~teamdriverstatus  = km~Teamdriverstatus
                                     AND kmbased~kmstart           = km~kmstart           AND kmbased~kmend             = km~kmend
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.
    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_kmbase.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_kmbase.

    MOVE-CORRESPONDING gt_km_based TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-Conditionunit = 'KM'.
      <fs_create>-%control      = VALUE #( Salesorganization = '01'
                                           Customer          = '01'
                                           Operationtype     = '01'
                                           Departuredistrict = '01'
                                           Departureprovince = '01'
                                           Originpoint       = '01'
                                           Arrivaldistrict   = '01'
                                           Arrivalprovince   = '01'
                                           Destinationpoint  = '01'
                                           Routetype         = '01'
                                           Teamdriverstatus  = '01'
                                           Vehiclebodytype   = '01'
                                           Vehicletype       = '01'
                                           Temperatureregime = '01'
                                           Kmstart           = '01'
                                           Kmend             = '01'
                                           Conditionprice    = '01'
                                           Conditionquantity = '01'
                                           Conditionunit     = '01'
                                           Currency          = '01'
                                           Validfrom         = '01'
                                           Validtodate       = '01' ).

      DATA(lv_old_kmend) = create_table[ Salesorganization = <fs_create>-Salesorganization
                                         customer          = <fs_create>-Customer
                                         operationtype     = <fs_create>-Operationtype
                                         departuredistrict = <fs_create>-Departuredistrict
                                         departureprovince = <fs_create>-Departureprovince
                                         originpoint       = <fs_create>-Originpoint
                                         arrivaldistrict   = <fs_create>-Arrivaldistrict
                                         arrivalprovince   = <fs_create>-Arrivalprovince
                                         destinationpoint  = <fs_create>-Destinationpoint
                                         routetype         = <fs_create>-Routetype
                                         vehiclebodytype   = <fs_create>-Vehiclebodytype
                                         vehicletype       = <fs_create>-Vehicletype
                                         temperatureregime = <fs_create>-Temperatureregime
                                         teamdriverstatus  = <fs_create>-Teamdriverstatus ]-kmend.

      IF lv_old_kmend = <fs_create>-Kmstart AND <fs_create>-Kmstart <> <fs_create>-Kmend.
        APPEND VALUE #( flag  = '5'
                        tabix = sy-tabix ) TO et_response.
      ENDIF.

    ENDLOOP.

    IF lt_data[] IS INITIAL.
      IF et_response IS INITIAL.
        MODIFY ENTITIES OF zr_sd_t_kmbase
               ENTITY ZrSdTKmbase CREATE AUTO FILL CID WITH create_table
               FAILED DATA(failed).
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ELSE.
      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        READ TABLE gt_km_based INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                            customer          = <fs_update>-Customer
                                                            operationtype     = <fs_update>-Operationtype
                                                            departuredistrict = <fs_update>-Departuredistrict
                                                            departureprovince = <fs_update>-Departureprovince
                                                            originpoint       = <fs_update>-Originpoint
                                                            arrivaldistrict   = <fs_update>-Arrivaldistrict
                                                            arrivalprovince   = <fs_update>-Arrivalprovince
                                                            destinationpoint  = <fs_update>-Destinationpoint
                                                            routetype         = <fs_update>-Routetype
                                                            vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                            vehicletype       = <fs_update>-Vehicletype
                                                            temperatureregime = <fs_update>-Temperatureregime
                                                            teamdriverstatus  = <fs_update>-Teamdriverstatus
                                                            kmstart           = <fs_update>-Kmstart
                                                            kmend             = <fs_update>-kmend.

        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-Validtodate.
        ENDIF.

        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_kmbase
               ENTITY ZrSdTKmbase UPDATE FIELDS ( Validfrom Validtodate ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_kmbase
               ENTITY ZrSdTKmbase CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_pallet_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'V' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_pallet_based[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_pallet_based WHERE customer = ''.

    LOOP AT gt_pallet_based ASSIGNING FIELD-SYMBOL(<fs_pallet_based>).
      DATA(lv_tabix) = sy-tabix.
                  IF <fs_pallet_based>-teamdriverstatus <> 'X'.
        <fs_pallet_based>-teamdriverstatus = ''.
      ENDIF.
      LOOP AT gt_pallet_based TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_pallet_based>-customer          AND operationtype     = <fs_pallet_based>-operationtype
                                                                             AND departureprovince = <fs_pallet_based>-departureprovince AND departuredistrict = <fs_pallet_based>-departuredistrict
                                                                             AND originpoint       = <fs_pallet_based>-originpoint       AND arrivaldistrict   = <fs_pallet_based>-arrivaldistrict
                                                                             AND arrivalprovince   = <fs_pallet_based>-arrivalprovince   AND destinationpoint  = <fs_pallet_based>-destinationpoint
                                                                             AND routetype         = <fs_pallet_based>-routetype         AND vehicletype       = <fs_pallet_based>-vehicletype
                                                                             AND vehiclebodytype   = <fs_pallet_based>-vehiclebodytype   AND temperatureregime = <fs_pallet_based>-temperatureregime
                                                                             AND teamdriverstatus  = <fs_pallet_based>-teamdriverstatus  AND packettype        = <fs_pallet_based>-packettype
                                                                             AND salesorganization = <fs_pallet_based>-salesorganization
                                                                             AND ( (     validfrom   = <fs_pallet_based>-validfrom
                                                                                     AND validtodate = <fs_pallet_based>-validtodate ) OR ( validtodate > <fs_pallet_based>-validfrom ) )
                                                                             AND ( quantityend = <fs_pallet_based>-quantitystart OR quantitystart = <fs_pallet_based>-quantitystart ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.
      <fs_pallet_based>-customer = |{ <fs_pallet_based>-customer ALPHA = IN }|.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '03'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_pallet_based ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_palletbsd.

    IF gt_pallet_based IS NOT INITIAL.

      SELECT pallet~uuid,
             pallet~salesorganization,
             pallet~customer,
             pallet~operationtype,
             pallet~departureprovince,
             pallet~departuredistrict,
             pallet~originpoint,
             pallet~arrivalprovince,
             pallet~arrivaldistrict,
             pallet~destinationpoint,
             pallet~routetype,
             pallet~vehicletype,
             pallet~vehiclebodytype,
             pallet~temperatureregime,
             pallet~quantitystart,
             pallet~quantityend,
             pallet~teamdriverstatus,
             pallet~packettype,
             pallet~validfrom,
             pallet~Validtodate
        FROM zsd_palletbsd    AS pallet
             INNER JOIN
             @gt_pallet_based AS pallet_based ON  pallet~customer          = pallet_based~Customer          AND pallet~salesorganization = pallet_based~salesorganization
                                              AND pallet~operationtype     = pallet_based~Operationtype     AND pallet~departureprovince = pallet_based~Departureprovince
                                              AND pallet~departuredistrict = pallet_based~Departuredistrict AND pallet~originpoint       = pallet_based~Originpoint
                                              AND pallet~arrivalprovince   = pallet_based~Arrivalprovince   AND pallet~arrivaldistrict   = pallet_based~Arrivaldistrict
                                              AND pallet~destinationpoint  = pallet_based~Destinationpoint  AND pallet~routetype         = pallet_based~Routetype
                                              AND pallet~vehicletype       = pallet_based~Vehicletype       AND pallet~vehiclebodytype   = pallet_based~Vehiclebodytype
                                              AND pallet~temperatureregime = pallet_based~Temperatureregime AND pallet~teamdriverstatus  = pallet_based~Teamdriverstatus
                                              AND pallet~packettype        = pallet_based~packettype
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_palletbsd.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_palletbsd.

    MOVE-CORRESPONDING gt_pallet_based TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control = VALUE #( Salesorganization = '01'
                                      Customer          = '01'
                                      Operationtype     = '01'
                                      Departuredistrict = '01'
                                      Departureprovince = '01'
                                      Originpoint       = '01'
                                      Arrivaldistrict   = '01'
                                      Arrivalprovince   = '01'
                                      Destinationpoint  = '01'
                                      Routetype         = '01'
                                      Teamdriverstatus  = '01'
                                      Vehiclebodytype   = '01'
                                      Vehicletype       = '01'
                                      Temperatureregime = '01'
                                      Quantitystart     = '01'
                                      Quantityend       = '01'
                                      Packettype        = '01'
                                      Conditionprice    = '01'
                                      Conditionquantity = '01'
                                      Currency          = '01'
                                      Validfrom         = '01'
                                      Validtodate       = '01' ).
      DATA(lv_old_quantityend) = create_table[ Salesorganization = <fs_create>-Salesorganization
                                               customer          = <fs_create>-Customer
                                               operationtype     = <fs_create>-Operationtype
                                               departuredistrict = <fs_create>-Departuredistrict
                                               departureprovince = <fs_create>-Departureprovince
                                               originpoint       = <fs_create>-Originpoint
                                               arrivaldistrict   = <fs_create>-Arrivaldistrict
                                               arrivalprovince   = <fs_create>-Arrivalprovince
                                               destinationpoint  = <fs_create>-Destinationpoint
                                               routetype         = <fs_create>-Routetype
                                               vehiclebodytype   = <fs_create>-Vehiclebodytype
                                               vehicletype       = <fs_create>-Vehicletype
                                               Packettype        = <fs_create>-Packettype
                                               temperatureregime = <fs_create>-Temperatureregime
                                               teamdriverstatus  = <fs_create>-Teamdriverstatus ]-Quantityend.
      IF lv_old_quantityend = <fs_create>-Quantitystart.
        APPEND VALUE #( flag  = '6'
                        tabix = sy-tabix ) TO et_response.
      ENDIF.
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_palletbsd
             ENTITY ZrSdPalletbsd CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ELSE.
      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).

        READ TABLE gt_pallet_based INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                                customer          = <fs_update>-Customer
                                                                operationtype     = <fs_update>-Operationtype
                                                                departuredistrict = <fs_update>-Departuredistrict
                                                                departureprovince = <fs_update>-Departureprovince
                                                                originpoint       = <fs_update>-Originpoint
                                                                arrivaldistrict   = <fs_update>-Arrivaldistrict
                                                                arrivalprovince   = <fs_update>-Arrivalprovince
                                                                destinationpoint  = <fs_update>-Destinationpoint
                                                                routetype         = <fs_update>-Routetype
                                                                vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                                vehicletype       = <fs_update>-Vehicletype
                                                                Packettype        = <fs_update>-Packettype
                                                                temperatureregime = <fs_update>-Temperatureregime
                                                                teamdriverstatus  = <fs_update>-Teamdriverstatus.
        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-ValidToDate.
        ENDIF.

        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_palletbsd
               ENTITY ZrSdPalletbsd UPDATE FIELDS ( Validfrom Validtodate ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_palletbsd
               ENTITY ZrSdPalletbsd CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_pointdiff_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'U' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_pointdiff[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_pointdiff WHERE customer = ''.

    LOOP AT gt_pointdiff ASSIGNING FIELD-SYMBOL(<fs_pointdiff>).
      DATA(lv_tabix) = sy-tabix.
      LOOP AT gt_pointdiff TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_pointdiff>-customer          AND operationtype     = <fs_pointdiff>-operationtype
                                                                          AND departureprovince = <fs_pointdiff>-departureprovince AND departuredistrict = <fs_pointdiff>-departuredistrict
                                                                          AND originpoint       = <fs_pointdiff>-originpoint       AND arrivaldistrict   = <fs_pointdiff>-arrivaldistrict
                                                                          AND arrivalprovince   = <fs_pointdiff>-arrivalprovince   AND destinationpoint  = <fs_pointdiff>-destinationpoint
                                                                          AND routetype         = <fs_pointdiff>-routetype         AND vehicletype       = <fs_pointdiff>-vehicletype
                                                                          AND vehiclebodytype   = <fs_pointdiff>-vehiclebodytype   AND temperatureregime = <fs_pointdiff>-temperatureregime
                                                                          AND salesorganization = <fs_pointdiff>-salesorganization
                                                                          AND ( (     validfrom = <fs_pointdiff>-validfrom
                                                                                  AND validto   = <fs_pointdiff>-validto ) OR ( validto > <fs_pointdiff>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_pointdiff>-customer = |{ <fs_pointdiff>-customer ALPHA = IN }|.
      IF <fs_pointdiff>-conditionunit = 'HB'.
        <fs_pointdiff>-conditionunit = 'EA'.
      ENDIF.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '07'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_pointdiff ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_t_pointdiff.

    IF gt_pointdiff IS NOT INITIAL.

      SELECT pointdiff~client,
             pointdiff~uuid,
             pointdiff~salesorganization,
             pointdiff~customer,
             pointdiff~operationtype,
             pointdiff~departureprovince,
             pointdiff~departuredistrict,
             pointdiff~originpoint,
             pointdiff~arrivalprovince,
             pointdiff~arrivaldistrict,
             pointdiff~destinationpoint,
             pointdiff~routecode,
             pointdiff~routetype,
             pointdiff~vehicletype,
             pointdiff~vehiclebodytype,
             pointdiff~temperatureregime,
             pointdiff~stopright,
             pointdiff~conditionquantity,
             pointdiff~conditionunit,
             pointdiff~conditionprice,
             pointdiff~currency,
             pointdiff~validfrom,
             pointdiff~validto
        FROM zsd_t_pointdiff AS pointdiff
             INNER JOIN
             @gt_pointdiff   AS point     ON  pointdiff~customer          = point~Customer          AND pointdiff~operationtype     = point~Operationtype
                                          AND pointdiff~departureprovince = point~departureprovince AND pointdiff~departuredistrict = point~departuredistrict
                                          AND pointdiff~originpoint       = point~Originpoint       AND pointdiff~arrivalprovince   = point~arrivalprovince
                                          AND pointdiff~arrivaldistrict   = point~arrivaldistrict   AND pointdiff~destinationpoint  = point~destinationpoint
                                          AND pointdiff~routecode         = point~routecode         AND pointdiff~routetype         = point~Routetype
                                          AND pointdiff~vehicletype       = point~Vehicletype       AND pointdiff~vehiclebodytype   = point~Vehiclebodytype
                                          AND pointdiff~temperatureregime = point~Temperatureregime AND pointdiff~salesorganization = point~salesorganization
        INTO TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_pointdiff.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_pointdiff.

    MOVE-CORRESPONDING gt_pointdiff TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control  = VALUE #( Salesorganization = '01'
                                       Customer          = '01'
                                       Operationtype     = '01'
                                       Departuredistrict = '01'
                                       Departureprovince = '01'
                                       Originpoint       = '01'
                                       Arrivaldistrict   = '01'
                                       Arrivalprovince   = '01'
                                       Destinationpoint  = '01'
                                       Routecode         = '01'
                                       Routetype         = '01'
                                       Vehiclebodytype   = '01'
                                       Vehicletype       = '01'
                                       Temperatureregime = '01'
                                       Stopright         = '01'
                                       Conditionprice    = '01'
                                       Conditionquantity = '01'
                                       Conditionunit     = '01'
                                       Currency          = '01'
                                       Validfrom         = '01'
                                       Validto           = '01' ).
      <fs_create>-Customer  = |{ <fs_create>-Customer ALPHA = IN } |.
      <fs_create>-Stopright = gt_pointdiff[ salesorganization = <fs_create>-Salesorganization
                                            customer          = <fs_create>-Customer
                                            operationtype     = <fs_create>-Operationtype
                                            departuredistrict = <fs_create>-Departuredistrict
                                            departureprovince = <fs_create>-Departureprovince
                                            destinationpoint  = <fs_create>-Destinationpoint
                                            arrivaldistrict   = <fs_create>-Arrivaldistrict
                                            arrivalprovince   = <fs_create>-Arrivalprovince
                                            originpoint       = <fs_create>-Originpoint
                                            routecode         = <fs_create>-Routecode
                                            routetype         = <fs_create>-Routetype
                                            vehiclebodytype   = <fs_create>-Vehiclebodytype
                                            vehicletype       = <fs_create>-Vehicletype ]-stopdifference.
    ENDLOOP.
    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_t_pointdiff
             ENTITY PointDifference CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      APPEND VALUE #( flag = '1' ) TO et_response.
    ELSE.
      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        READ TABLE gt_pointdiff INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                             customer          = <fs_update>-Customer
                                                             operationtype     = <fs_update>-Operationtype
                                                             departuredistrict = <fs_update>-Departuredistrict
                                                             departureprovince = <fs_update>-Departureprovince
                                                             originpoint       = <fs_update>-Originpoint
                                                             arrivaldistrict   = <fs_update>-Arrivaldistrict
                                                             arrivalprovince   = <fs_update>-Arrivalprovince
                                                             destinationpoint  = <fs_update>-Destinationpoint
                                                             routetype         = <fs_update>-Routetype
                                                             vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                             vehicletype       = <fs_update>-Vehicletype
                                                             temperatureregime = <fs_update>-Temperatureregime.
        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-validto.
        ENDIF.
        IF lv_date_old_start <= <fs_update>-Validto AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validto = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validto AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_pointdiff
               ENTITY PointDifference UPDATE FIELDS ( Validfrom Validto ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_pointdiff
               ENTITY PointDifference CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_porterage_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).
    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'V' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_porterage[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_porterage WHERE customer = ''.

    LOOP AT gt_porterage ASSIGNING FIELD-SYMBOL(<fs_porterage>).
      DATA(lv_tabix) = sy-tabix.
      LOOP AT gt_porterage TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_porterage>-customer          AND operationtype     = <fs_porterage>-operationtype
                                                                          AND departureprovince = <fs_porterage>-departureprovince AND departuredistrict = <fs_porterage>-departuredistrict
                                                                          AND originpoint       = <fs_porterage>-originpoint       AND arrivaldistrict   = <fs_porterage>-arrivaldistrict
                                                                          AND arrivalprovince   = <fs_porterage>-arrivalprovince   AND destinationpoint  = <fs_porterage>-destinationpoint
                                                                          AND routetype         = <fs_porterage>-routetype         AND vehicletype       = <fs_porterage>-vehicletype
                                                                          AND vehiclebodytype   = <fs_porterage>-vehiclebodytype   AND temperatureregime = <fs_porterage>-temperatureregime
                                                                          AND quantitystart     = <fs_porterage>-quantitystart     AND quantityend       = <fs_porterage>-quantityend
                                                                          AND salesorganization = <fs_porterage>-salesorganization
                                                                          AND ( (     validfrom   = <fs_porterage>-validfrom
                                                                                  AND validtodate = <fs_porterage>-validtodate ) OR ( validtodate > <fs_porterage>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_porterage>-customer = |{ <fs_porterage>-customer ALPHA = IN }|.
      IF <fs_porterage>-multipledrivercond <> 'X'.
        <fs_porterage>-multipledrivercond = ''.
      ENDIF.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '09'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_porterage ).
    IF lt_response IS NOT INITIAL.
      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_t_porterage.
    IF gt_porterage IS NOT INITIAL.

      SELECT porterage~client,
             porterage~uuid,
             porterage~salesorganization,
             porterage~customer,
             porterage~operationtype,
             porterage~departureprovince,
             porterage~departuredistrict,
             porterage~originpoint,
             porterage~arrivalprovince,
             porterage~arrivaldistrict,
             porterage~destinationpoint,
             porterage~routetype,
             porterage~multipledrivercond,
             porterage~vehicletype,
             porterage~vehiclebodytype,
             porterage~temperatureregime,
             porterage~quantitystart,
             porterage~quantityend,
             porterage~conditionquantity,
             porterage~conditionunit,
             porterage~conditionprice,
             porterage~currency,
             porterage~validfrom,
             porterage~validtodate
        FROM zsd_t_porterage AS porterage
             INNER JOIN
             @gt_porterage   AS p         ON  porterage~customer          = p~Customer          AND porterage~operationtype      = p~Operationtype
                                          AND porterage~arrivalprovince   = p~Arrivalprovince   AND porterage~arrivaldistrict    = p~Arrivaldistrict
                                          AND porterage~originpoint       = p~Originpoint       AND porterage~departuredistrict  = p~Departuredistrict
                                          AND porterage~destinationpoint  = p~Destinationpoint  AND porterage~routetype          = p~Routetype
                                          AND porterage~vehicletype       = p~Vehicletype       AND porterage~vehiclebodytype    = p~Vehiclebodytype
                                          AND porterage~temperatureregime = p~Temperatureregime AND porterage~multipledrivercond = p~Multipledrivercond
                                          AND porterage~salesorganization = p~salesorganization
                                          AND ( porterage~quantitystart = p~Quantitystart AND porterage~quantityend = p~Quantityend )
        INTO TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_porterage.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_porterage.

    MOVE-CORRESPONDING gt_porterage TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control = VALUE #( Salesorganization  = '01'
                                      Customer           = '01'
                                      Operationtype      = '01'
                                      Departuredistrict  = '01'
                                      Departureprovince  = '01'
                                      Originpoint        = '01'
                                      Arrivaldistrict    = '01'
                                      Arrivalprovince    = '01'
                                      Destinationpoint   = '01'
                                      Routetype          = '01'
                                      Multipledrivercond = '01'
                                      Vehiclebodytype    = '01'
                                      Vehicletype        = '01'
                                      Temperatureregime  = '01'
                                      Quantitystart      = '01'
                                      Quantityend        = '01'
                                      Conditionprice     = '01'
                                      Conditionquantity  = '01'
                                      Conditionunit      = '01'
                                      Currency           = '01'
                                      Validfrom          = '01'
                                      Validtodate        = '01' ).
      DATA(lv_old_Quantityend) = create_table[ salesorganization = <fs_create>-Salesorganization
                                               customer          = <fs_create>-Customer
                                               operationtype     = <fs_create>-Operationtype
                                               departuredistrict = <fs_create>-Departuredistrict
                                               departureprovince = <fs_create>-Departureprovince
                                               originpoint       = <fs_create>-Originpoint
                                               arrivaldistrict   = <fs_create>-Arrivaldistrict
                                               arrivalprovince   = <fs_create>-Arrivalprovince
                                               destinationpoint  = <fs_create>-Destinationpoint
                                               routetype         = <fs_create>-Routetype
                                               vehiclebodytype   = <fs_create>-Vehiclebodytype
                                               vehicletype       = <fs_create>-Vehicletype
                                               temperatureregime = <fs_create>-Temperatureregime ]-Quantityend.

      IF lv_old_Quantityend = <fs_create>-Quantitystart.
        APPEND VALUE #( flag  = '6'
                        tabix = sy-tabix ) TO et_response.
      ENDIF.
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_t_porterage
             ENTITY Porterage CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ELSE.

      MOVE-CORRESPONDING lt_data TO update_table.
      DATA(lv_tabindex) = 1.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).

        READ TABLE gt_porterage INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                             customer          = <fs_update>-Customer
                                                             operationtype     = <fs_update>-Operationtype
                                                             departuredistrict = <fs_update>-Departuredistrict
                                                             departureprovince = <fs_update>-Departureprovince
                                                             originpoint       = <fs_update>-Originpoint
                                                             arrivaldistrict   = <fs_update>-Arrivaldistrict
                                                             arrivalprovince   = <fs_update>-Arrivalprovince
                                                             destinationpoint  = <fs_update>-Destinationpoint
                                                             routetype         = <fs_update>-Routetype
                                                             vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                             vehicletype       = <fs_update>-Vehicletype
                                                             temperatureregime = <fs_update>-Temperatureregime
                                                             quantitystart     = <fs_update>-Quantitystart
                                                             quantityend       = <fs_update>-Quantityend.
        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-ValidToDate.
        ELSE.
          DELETE update_table INDEX lv_tabindex.        "#EC CI_NOORDER
          CONTINUE.
        ENDIF.

        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX lv_tabindex.        "#EC CI_NOORDER
        ENDIF.
        lv_tabindex += 1.
      ENDLOOP.
      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_porterage
               ENTITY Porterage UPDATE FIELDS ( Validfrom Validtodate ) WITH update_table.
      ENDIF.
      MODIFY ENTITIES OF zr_sd_t_porterage
             ENTITY Porterage CREATE AUTO FILL CID WITH create_table
             FAILED failed.
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_route_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'R' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_route_based[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_route_based WHERE customer = ''.

    LOOP AT gt_route_based ASSIGNING FIELD-SYMBOL(<fs_route_based>).
      DATA(lv_tabix) = sy-tabix.
       IF <fs_route_based>-teamdriversts <> 'X'.
        <fs_route_based>-teamdriversts = ''.
      ENDIF.
      LOOP AT gt_route_based TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_route_based>-customer          AND operationtype     = <fs_route_based>-operationtype
                                                                            AND departureprovince = <fs_route_based>-departureprovince AND departuredistrict = <fs_route_based>-departuredistrict
                                                                            AND originpoint       = <fs_route_based>-originpoint       AND routetype         = <fs_route_based>-routetype
                                                                            AND vehicletype       = <fs_route_based>-vehicletype       AND vehiclebodytype   = <fs_route_based>-vehiclebodytype
                                                                            AND temperatureregime = <fs_route_based>-temperatureregime AND teamdriversts     = <fs_route_based>-teamdriversts
                                                                            AND routecode         = <fs_route_based>-routecode         AND salesorganization = <fs_route_based>-salesorganization
                                                                            AND ( (     validfrom = <fs_route_based>-validfrom
                                                                                    AND validto   = <fs_route_based>-validto ) OR ( validto > <fs_route_based>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_route_based>-customer = |{ <fs_route_based>-customer ALPHA = IN }|.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '05'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_route_based ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_t_routebsd.

    IF gt_route_based IS NOT INITIAL.

      SELECT routebsd~client,
             routebsd~uuid,
             routebsd~salesorganization,
             routebsd~customer,
             routebsd~operationtype,
             routebsd~departureprovince,
             routebsd~departuredistrict,
             routebsd~originpoint,
             routebsd~routecode,
             routebsd~routetype,
             routebsd~teamdriversts,
             routebsd~vehicletype,
             routebsd~vehiclebodytype,
             routebsd~temperatureregime,
             routebsd~conditionquantity,
             routebsd~conditionunit,
             routebsd~conditionprice,
             routebsd~currency,
             routebsd~validfrom,
             routebsd~validto
        FROM zsd_t_routebsd  AS routebsd
             INNER JOIN
             @gt_route_based AS route    ON  routebsd~customer          = route~Customer          AND routebsd~operationtype     = route~Operationtype
                                         AND routebsd~departureprovince = route~departureprovince AND routebsd~departuredistrict = route~departuredistrict
                                         AND routebsd~originpoint       = route~Originpoint       AND routebsd~routecode         = route~routecode
                                         AND routebsd~routetype         = route~Routetype         AND routebsd~teamdriversts     = route~teamdriversts
                                         AND routebsd~vehicletype       = route~Vehicletype       AND routebsd~vehiclebodytype   = route~Vehiclebodytype
                                         AND routebsd~temperatureregime = route~Temperatureregime AND routebsd~salesorganization = route~salesorganization
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_routebsd.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_routebsd.

    MOVE-CORRESPONDING gt_route_based TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).

      <fs_create>-%control = VALUE #( Salesorganization = '01'
                                      Customer          = '01'
                                      Operationtype     = '01'
                                      Departuredistrict = '01'
                                      Departureprovince = '01'
                                      Originpoint       = '01'
                                      Routecode         = '01'
                                      Routetype         = '01'
                                      Teamdriversts     = '01'
                                      Vehiclebodytype   = '01'
                                      Vehicletype       = '01'
                                      Temperatureregime = '01'
                                      Conditionprice    = '01'
                                      Conditionquantity = '01'
                                      Conditionunit     = '01'
                                      Currency          = '01'
                                      Validfrom         = '01'
                                      Validto           = '01' ).
      <fs_create>-Customer = |{ <fs_create>-Customer ALPHA = IN } |.
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_t_routebsd
             ENTITY RouteBased CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      APPEND VALUE #( flag = '1' ) TO et_response.
    ELSE.

      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        READ TABLE gt_route_based INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                               customer          = <fs_update>-Customer
                                                               operationtype     = <fs_update>-Operationtype
                                                               departuredistrict = <fs_update>-Departuredistrict
                                                               departureprovince = <fs_update>-Departureprovince
                                                               originpoint       = <fs_update>-Originpoint
                                                               routetype         = <fs_update>-Routetype
                                                               vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                               vehicletype       = <fs_update>-Vehicletype
                                                               temperatureregime = <fs_update>-Temperatureregime
                                                               teamdriversts     = <fs_update>-Teamdriversts.
        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-validto.
        ENDIF.

        IF lv_date_old_start <= <fs_update>-Validto AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validto = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validto AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_routebsd
               ENTITY RouteBased UPDATE FIELDS ( Validfrom Validto ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_routebsd
               ENTITY RouteBased CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_shipping_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'T' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_shipping_based[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_shipping_based WHERE customer = ''.

    LOOP AT gt_shipping_based ASSIGNING FIELD-SYMBOL(<fs_shipping_based>).
      DATA(lv_tabix) = sy-tabix.
      IF <fs_shipping_based>-teamdriverstatus <> 'X'.
        <fs_shipping_based>-teamdriverstatus = ''.
      ENDIF.
      LOOP AT gt_shipping_based TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_shipping_based>-customer          AND operationtype     = <fs_shipping_based>-operationtype
                                                                               AND departureprovince = <fs_shipping_based>-departureprovince AND departuredistrict = <fs_shipping_based>-departuredistrict
                                                                               AND originpoint       = <fs_shipping_based>-originpoint       AND arrivaldistrict   = <fs_shipping_based>-arrivaldistrict
                                                                               AND arrivalprovince   = <fs_shipping_based>-arrivalprovince   AND destinationpoint  = <fs_shipping_based>-destinationpoint
                                                                               AND routetype         = <fs_shipping_based>-routetype         AND vehicletype       = <fs_shipping_based>-vehicletype
                                                                               AND vehiclebodytype   = <fs_shipping_based>-vehiclebodytype   AND temperatureregime = <fs_shipping_based>-temperatureregime
                                                                               AND teamdriverstatus  = <fs_shipping_based>-teamdriverstatus  AND salesorganization = <fs_shipping_based>-salesorganization
                                                                               AND ( (     validfrom   = <fs_shipping_based>-validfrom
                                                                                       AND validtodate = <fs_shipping_based>-validtodate ) OR ( validtodate > <fs_shipping_based>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_shipping_based>-customer = |{ <fs_shipping_based>-customer ALPHA = IN }|.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '01'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_shipping_based ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_shippingbsd.

    IF gt_shipping_based IS NOT INITIAL.

      SELECT shippingbsd~uuid,
             shippingbsd~salesorganization,
             shippingbsd~customer,
             shippingbsd~operationtype,
             shippingbsd~departureprovince,
             shippingbsd~departuredistrict,
             shippingbsd~originpoint,
             shippingbsd~arrivalprovince,
             shippingbsd~arrivaldistrict,
             shippingbsd~destinationpoint,
             shippingbsd~routetype,
             shippingbsd~teamdriverstatus,
             shippingbsd~vehicletype,
             shippingbsd~vehiclebodytype,
             shippingbsd~temperatureregime,
             shippingbsd~conditionprice,
             shippingbsd~validfrom,
             shippingbsd~validtodate,
             shippingbsd~conditionquantity
        FROM zsd_shippingbsd    AS shippingbsd
             INNER JOIN
             @gt_shipping_based AS shipping    ON  shippingbsd~customer          = shipping~Customer          AND shippingbsd~salesorganization = shipping~salesorganization
                                               AND shippingbsd~operationtype     = shipping~Operationtype     AND shippingbsd~departureprovince = shipping~Departureprovince
                                               AND shippingbsd~departuredistrict = shipping~Departuredistrict AND shippingbsd~originpoint       = shipping~Originpoint
                                               AND shippingbsd~arrivalprovince   = shipping~Arrivalprovince   AND shippingbsd~arrivaldistrict   = shipping~Arrivaldistrict
                                               AND shippingbsd~destinationpoint  = shipping~Destinationpoint  AND shippingbsd~routetype         = shipping~Routetype
                                               AND shippingbsd~vehicletype       = shipping~Vehicletype       AND shippingbsd~vehiclebodytype   = shipping~Vehiclebodytype
                                               AND shippingbsd~temperatureregime = shipping~Temperatureregime AND shippingbsd~teamdriverstatus  = shipping~Teamdriverstatus
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_shippingbsd.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_shippingbsd.

    MOVE-CORRESPONDING gt_shipping_based TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).

      <fs_create>-%control = VALUE #( Salesorganization = '01'
                                      Arrivaldistrict   = '01'
                                      Arrivalprovince   = '01'
                                      Temperatureregime = '01'
                                      Vehiclebodytype   = '01'
                                      Vehicletype       = '01'
                                      Routetype         = '01'
                                      Destinationpoint  = '01'
                                      Originpoint       = '01'
                                      Departuredistrict = '01'
                                      Departureprovince = '01'
                                      Operationtype     = '01'
                                      customer          = '01'
                                      Validfrom         = '01'
                                      Validtodate       = '01'
                                      Conditionquantity = '01'
                                      Conditionunit     = '01'
                                      Conditionprice    = '01'
                                      Currency          = '01'
                                      Teamdriverstatus  = '01' ).
    ENDLOOP.
    IF lt_data[] IS INITIAL AND et_response IS INITIAL.

      MODIFY ENTITIES OF zr_sd_shippingbsd
             ENTITY ZrSdShippingbsd CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      APPEND VALUE #( flag = '1' ) TO et_response.
    ELSE.
      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        READ TABLE gt_shipping_based INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                                  customer          = <fs_update>-Customer
                                                                  operationtype     = <fs_update>-Operationtype
                                                                  departuredistrict = <fs_update>-Departuredistrict
                                                                  departureprovince = <fs_update>-Departureprovince
                                                                  originpoint       = <fs_update>-Originpoint
                                                                  arrivaldistrict   = <fs_update>-Arrivaldistrict
                                                                  arrivalprovince   = <fs_update>-Arrivalprovince
                                                                  destinationpoint  = <fs_update>-Destinationpoint
                                                                  routetype         = <fs_update>-Routetype
                                                                  vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                                  vehicletype       = <fs_update>-Vehicletype
                                                                  temperatureregime = <fs_update>-Temperatureregime
                                                                  teamdriverstatus  = <fs_update>-Teamdriverstatus.
        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-ValidToDate.
        ENDIF.

        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_shippingbsd
               ENTITY ZrSdShippingbsd UPDATE FIELDS ( validfrom Validtodate ) WITH update_table
               " TODO: variable is assigned but never used (ABAP cleaner)
               FAILED DATA(update_failed).
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_shippingbsd
               ENTITY ZrSdShippingbsd CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_tollcost_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'M' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_tollcost[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_tollcost WHERE customer = ''.

    LOOP AT gt_tollcost ASSIGNING FIELD-SYMBOL(<fs_tollcost>).
      DATA(lv_tabix) = sy-tabix.
      LOOP AT gt_tollcost TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer        = <fs_tollcost>-customer        AND vehicletype       = <fs_tollcost>-vehicletype
                                                                         AND vehiclebodytype = <fs_tollcost>-vehiclebodytype AND temperatureregime = <fs_tollcost>-temperatureregime
                                                                         AND routecode       = <fs_tollcost>-routecode       AND operationtype     = <fs_tollcost>-operationtype
                                                                         AND salesorganization = <fs_tollcost>-salesorganization
                                                                         AND ( (     validfrom   = <fs_tollcost>-validfrom
                                                                                 AND validtodate = <fs_tollcost>-validtodate ) OR ( validtodate > <fs_tollcost>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_tollcost>-customer      = |{ <fs_tollcost>-customer ALPHA = IN }|.
      <fs_tollcost>-conditionunit = 'SFR'.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '11'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_tollcost ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_t_tollcost.

    IF gt_tollcost IS NOT INITIAL.

      SELECT DISTINCT tollcost~client,
                      tollcost~uuid,
                      tollcost~salesorganization,
                      tollcost~customer,
                      tollcost~operationtype,
                      tollcost~vehicletype,
                      tollcost~vehiclebodytype,
                      tollcost~temperatureregime,
                      tollcost~routecode,
                      tollcost~conditionquantity,
                      tollcost~conditionunit,
                      tollcost~conditionprice,
                      tollcost~currency,
                      tollcost~validfrom,
                      tollcost~validtodate
        FROM zsd_t_tollcost AS tollcost
             INNER JOIN
             @gt_tollcost   AS toll     ON  tollcost~customer          = toll~Customer          AND tollcost~operationtype   = toll~operationtype
                                        AND tollcost~vehicletype       = toll~Vehicletype       AND tollcost~vehiclebodytype = toll~Vehiclebodytype
                                        AND tollcost~temperatureregime = toll~Temperatureregime AND tollcost~routecode       = toll~Routecode
                                        AND tollcost~salesorganization = toll~salesorganization
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_tollcost.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_tollcost.

    MOVE-CORRESPONDING gt_tollcost TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control = VALUE #( Salesorganization = '01'
                                      Customer          = '01'
                                      Operationtype     = '01'
                                      Vehiclebodytype   = '01'
                                      Vehicletype       = '01'
                                      Temperatureregime = '01'
                                      Routecode         = '01'
                                      Conditionprice    = '01'
                                      Conditionquantity = '01'
                                      Conditionunit     = '01'
                                      Currency          = '01'
                                      Validfrom         = '01'
                                      Validtodate       = '01' ).
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_t_tollcost
             ENTITY TollCost CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      APPEND VALUE #( flag = '1' ) TO et_response.
    ELSE.

      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        DATA(lv_date_old_start) = gt_tollcost[ salesorganization = <fs_update>-Salesorganization
                                               customer          = <fs_update>-Customer
                                               operationtype     = <fs_update>-Operationtype
                                               routecode         = <fs_update>-Routecode
                                               vehiclebodytype   = <fs_update>-Vehiclebodytype
                                               vehicletype       = <fs_update>-Vehicletype
                                               temperatureregime = <fs_update>-Temperatureregime ]-ValidFrom.

        DATA(lv_date_old_end) = gt_tollcost[ salesorganization = <fs_update>-Salesorganization
                                             customer          = <fs_update>-Customer
                                             operationtype     = <fs_update>-Operationtype
                                             routecode         = <fs_update>-Routecode
                                             vehiclebodytype   = <fs_update>-Vehiclebodytype
                                             vehicletype       = <fs_update>-Vehicletype
                                             temperatureregime = <fs_update>-Temperatureregime ]-ValidToDate.

        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_tollcost
               ENTITY TollCost UPDATE FIELDS ( Validfrom Validtodate ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_tollcost
               ENTITY TollCost CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD read_waiting_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'N' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_waiting[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_waiting WHERE customer IS INITIAL.

    LOOP AT gt_waiting ASSIGNING FIELD-SYMBOL(<fs_data>).
      <fs_data>-customer      = |{ <fs_data>-customer ALPHA = IN }|.
      <fs_data>-conditionunit = 'SFR'.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING
                                               iv_table_type     = '10'
                                             IMPORTING
                                               et_check_response = DATA(lt_response)
                                               CHANGING ct_check_fields   = gt_waiting ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.
    DATA lt_data TYPE TABLE OF zsd_t_waiting.

    IF gt_waiting IS NOT INITIAL.

      SELECT DISTINCT waiting~client,
                      waiting~uuid,
                      waiting~salesorganization,
                      waiting~customer,
                      waiting~operationtype,
                      waiting~vehicletype,
                      waiting~vehiclebodytype,
                      waiting~temperatureregime,
                      waiting~minwaiting,
                      waiting~conditionquantity,
                      waiting~conditionunit,
                      waiting~minconditionprice,
                      waiting~conditionprice,
                      waiting~currency,
                      waiting~validfrom,
                      waiting~validtodate
        FROM zsd_t_waiting AS waiting
             INNER JOIN
             @gt_waiting   AS wait    ON  waiting~customer      = wait~Customer      AND waiting~temperatureregime = wait~Temperatureregime
                                      AND waiting~operationtype = wait~operationtype AND waiting~salesorganization = wait~salesorganization
                                      AND waiting~vehicletype   = wait~Vehicletype   AND waiting~vehiclebodytype   = wait~Vehiclebodytype
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_waiting.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_waiting.

    MOVE-CORRESPONDING gt_waiting TO create_table.
    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control = VALUE #( Salesorganization = '01'
                                      Customer          = '01'
                                      Operationtype     = '01'
                                      Vehicletype       = '01'
                                      Vehiclebodytype   = '01'
                                      Temperatureregime = '01'
                                      Minwaiting        = '01'
                                      Conditionquantity = '01'
                                      Conditionunit     = '01'
                                      MinConditionprice = '01'
                                      Conditionprice    = '01'
                                      Currency          = '01'
                                      Validfrom         = '01'
                                      Validtodate       = '01' ).
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_t_waiting
             ENTITY WaitingTime CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ELSE.

      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        DATA(lv_date_old_start) = gt_waiting[ salesorganization = <fs_update>-Salesorganization
                                              customer          = <fs_update>-Customer
                                              operationtype     = <fs_update>-Operationtype
                                              vehiclebodytype   = <fs_update>-Vehiclebodytype
                                              vehicletype       = <fs_update>-Vehicletype
                                              temperatureregime = <fs_update>-Temperatureregime ]-ValidFrom.

        DATA(lv_date_old_end) = gt_waiting[ salesorganization = <fs_update>-Salesorganization
                                            customer          = <fs_update>-Customer
                                            operationtype     = <fs_update>-Operationtype
                                            vehiclebodytype   = <fs_update>-Vehiclebodytype
                                            vehicletype       = <fs_update>-Vehicletype
                                            temperatureregime = <fs_update>-Temperatureregime ]-ValidToDate.

        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_waiting
               ENTITY WaitingTime UPDATE FIELDS ( Validfrom Validtodate ) WITH update_table
               " TODO: variable is assigned but never used (ABAP cleaner)
               MAPPED DATA(mapped)
               FAILED failed
               " TODO: variable is assigned but never used (ABAP cleaner)
               REPORTED DATA(reported).
      ENDIF.

      MODIFY ENTITIES OF zr_sd_t_waiting
             ENTITY WaitingTime CREATE AUTO FILL CID WITH create_table
             MAPPED mapped
             REPORTED reported
             FAILED failed.
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.

    ENDIF.
  ENDMETHOD.


  METHOD read_weight_excel.
    DATA(lo_read_access) = xco_cp_xlsx=>document->for_file_content( attachment )->read_access( ).

    DATA(lo_first_worksheet) = lo_read_access->get_workbook(
    )->worksheet->at_position( 1 ).

    DATA(lo_pattern_1) = xco_cp_xlsx_selection=>pattern_builder->simple_from_to(
      )->from_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'A' )
      )->to_column( xco_cp_xlsx=>coordinate->for_alphabetic_value( 'V' )
      )->from_row( xco_cp_xlsx=>coordinate->for_numeric_value( 2 )
      )->get_pattern( ).

    lo_first_worksheet->select( lo_pattern_1
      )->row_stream(
      )->operation->write_to( REF #( gt_weight_based[] )
      )->if_xco_xlsx_ra_operation~execute( ).
    DELETE gt_weight_based WHERE customer = ''.

    LOOP AT gt_weight_based ASSIGNING FIELD-SYMBOL(<fs_weight_based>).
      DATA(lv_tabix) = sy-tabix.
      IF <fs_weight_based>-teamdriverstatus <> 'X'.
        <fs_weight_based>-teamdriverstatus = ''.
      ENDIF.
      LOOP AT gt_weight_based TRANSPORTING NO FIELDS FROM lv_tabix + 1 WHERE     customer          = <fs_weight_based>-customer          AND operationtype     = <fs_weight_based>-operationtype
                                                                             AND departureprovince = <fs_weight_based>-departureprovince AND departuredistrict = <fs_weight_based>-departuredistrict
                                                                             AND originpoint       = <fs_weight_based>-originpoint       AND arrivaldistrict   = <fs_weight_based>-arrivaldistrict
                                                                             AND arrivalprovince   = <fs_weight_based>-arrivalprovince   AND destinationpoint  = <fs_weight_based>-destinationpoint
                                                                             AND routetype         = <fs_weight_based>-routetype         AND vehicletype       = <fs_weight_based>-vehicletype
                                                                             AND vehiclebodytype   = <fs_weight_based>-vehiclebodytype   AND temperatureregime = <fs_weight_based>-temperatureregime
                                                                             AND teamdriverstatus  = <fs_weight_based>-teamdriverstatus  AND quantitystart     = <fs_weight_based>-quantitystart
                                                                             AND quantityend       = <fs_weight_based>-quantityend       AND salesorganization = <fs_weight_based>-salesorganization
                                                                             AND ( (     validfrom   = <fs_weight_based>-validfrom
                                                                                     AND validtodate = <fs_weight_based>-validtodate ) OR ( validtodate > <fs_weight_based>-validfrom ) ).
        EXIT.
      ENDLOOP.

      IF sy-subrc = 0.
        APPEND VALUE #( flag  = '4'
                        tabix = lv_tabix + 1 ) TO et_response.
        RETURN.
      ENDIF.

      <fs_weight_based>-customer      = |{ <fs_weight_based>-customer ALPHA = IN }|.
      <fs_weight_based>-conditionunit = 'KG'.
    ENDLOOP.

    zcl_sd_fields_validations=>check_fields( EXPORTING iv_table_type     = '04'
                                             IMPORTING et_check_response = DATA(lt_response)
                                             CHANGING  ct_check_fields   = gt_weight_based ).

    IF lt_response IS NOT INITIAL.

      MOVE-CORRESPONDING lt_response TO et_response.
      RETURN.
    ENDIF.

    DATA lt_data TYPE TABLE OF zsd_t_weightbsd.

    IF gt_weight_based IS NOT INITIAL.

      SELECT DISTINCT weight~uuid,
                      weight~customer,
                      weight~salesorganization,
                      weight~operationtype,
                      weight~departureprovince,
                      weight~departuredistrict,
                      weight~originpoint,
                      weight~arrivalprovince,
                      weight~arrivaldistrict,
                      weight~destinationpoint,
                      weight~routetype,
                      weight~vehicletype,
                      weight~vehiclebodytype,
                      weight~temperatureregime,
                      weight~teamdriverstatus,
                      weight~validfrom,
                      weight~validtodate,
                      weight~quantitystart,
                      weight~quantityend
        FROM zsd_t_weightbsd AS weight
               INNER JOIN
                 @gt_weight_based AS weight_based ON  weight~customer          = weight_based~Customer          AND weight~salesorganization = weight_based~salesorganization
                                                  AND weight~operationtype     = weight_based~Operationtype     AND weight~departureprovince = weight_based~Departureprovince
                                                  AND weight~departuredistrict = weight_based~Departuredistrict AND weight~originpoint       = weight_based~Originpoint
                                                  AND weight~arrivalprovince   = weight_based~Arrivalprovince   AND weight~arrivaldistrict   = weight_based~Arrivaldistrict
                                                  AND weight~destinationpoint  = weight_based~Destinationpoint  AND weight~routetype         = weight_based~Routetype
                                                  AND weight~vehicletype       = weight_based~Vehicletype       AND weight~vehiclebodytype   = weight_based~Vehiclebodytype
                                                  AND weight~temperatureregime = weight_based~Temperatureregime AND weight~quantitystart     = weight_based~Quantitystart
                                                  AND weight~quantityend       = weight_based~Quantityend       AND weight~teamdriverstatus  = weight_based~Teamdriverstatus
        INTO CORRESPONDING FIELDS OF TABLE @lt_data.
    ENDIF.

    DATA create_table TYPE TABLE FOR CREATE zr_sd_t_weightbsd.
    DATA update_table TYPE TABLE FOR UPDATE zr_sd_t_weightbsd.

    MOVE-CORRESPONDING gt_weight_based TO create_table.

    LOOP AT create_table ASSIGNING FIELD-SYMBOL(<fs_create>).
      <fs_create>-%control = VALUE #( Salesorganization = '01'
                                      Customer          = '01'
                                      Operationtype     = '01'
                                      Departuredistrict = '01'
                                      Departureprovince = '01'
                                      Originpoint       = '01'
                                      Arrivaldistrict   = '01'
                                      Arrivalprovince   = '01'
                                      Destinationpoint  = '01'
                                      Routetype         = '01'
                                      Teamdriverstatus  = '01'
                                      Vehiclebodytype   = '01'
                                      Vehicletype       = '01'
                                      Temperatureregime = '01'
                                      Quantityend       = '01'
                                      Quantitystart     = '01'
                                      Conditionprice    = '01'
                                      Conditionquantity = '01'
                                      Conditionunit     = '01'
                                      Currency          = '01'
                                      Validfrom         = '01'
                                      Validtodate       = '01' ).
      DATA(lv_old_Quantityend) = create_table[ salesorganization = <fs_create>-Salesorganization
                                               customer          = <fs_create>-Customer
                                               operationtype     = <fs_create>-Operationtype
                                               departuredistrict = <fs_create>-Departuredistrict
                                               departureprovince = <fs_create>-Departureprovince
                                               originpoint       = <fs_create>-Originpoint
                                               arrivaldistrict   = <fs_create>-Arrivaldistrict
                                               arrivalprovince   = <fs_create>-Arrivalprovince
                                               destinationpoint  = <fs_create>-Destinationpoint
                                               routetype         = <fs_create>-Routetype
                                               vehiclebodytype   = <fs_create>-Vehiclebodytype
                                               vehicletype       = <fs_create>-Vehicletype
                                               temperatureregime = <fs_create>-Temperatureregime
                                               teamdriverstatus  = <fs_create>-Teamdriverstatus ]-Quantityend.

      IF lv_old_Quantityend = <fs_create>-quantitystart.
        APPEND VALUE #( flag  = '6'
                        tabix = sy-tabix ) TO et_response.
      ENDIF.
    ENDLOOP.

    IF lt_data[] IS INITIAL.

      MODIFY ENTITIES OF zr_sd_t_weightbsd
             ENTITY ZrSdTWeightbsd CREATE AUTO FILL CID WITH create_table
             FAILED DATA(failed).
      IF failed IS INITIAL.
        APPEND VALUE #( flag = '1' ) TO et_response.
      ENDIF.
    ELSE.
      MOVE-CORRESPONDING lt_data TO update_table.

      LOOP AT update_table ASSIGNING FIELD-SYMBOL(<fs_update>).
        READ TABLE gt_weight_based INTO DATA(ls_check) WITH KEY salesorganization = <fs_update>-Salesorganization
                                                                customer          = <fs_update>-Customer
                                                                operationtype     = <fs_update>-Operationtype
                                                                departuredistrict = <fs_update>-Departuredistrict
                                                                departureprovince = <fs_update>-Departureprovince
                                                                originpoint       = <fs_update>-Originpoint
                                                                arrivaldistrict   = <fs_update>-Arrivaldistrict
                                                                arrivalprovince   = <fs_update>-Arrivalprovince
                                                                destinationpoint  = <fs_update>-Destinationpoint
                                                                routetype         = <fs_update>-Routetype
                                                                vehiclebodytype   = <fs_update>-Vehiclebodytype
                                                                vehicletype       = <fs_update>-Vehicletype
                                                                temperatureregime = <fs_update>-Temperatureregime
                                                                teamdriverstatus  = <fs_update>-Teamdriverstatus
                                                                quantitystart     = <fs_update>-Quantitystart
                                                                quantityend       = <fs_update>-Quantityend.
        IF ls_check IS NOT INITIAL.
          DATA(lv_date_old_start) = ls_check-ValidFrom.
          DATA(lv_date_old_end) = ls_check-ValidToDate.
        ENDIF.

        IF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start >= <fs_update>-Validfrom.
          <fs_update>-Validtodate = lv_date_old_start - 1.
        ELSEIF lv_date_old_start <= <fs_update>-Validtodate AND lv_date_old_end >= <fs_update>-Validfrom AND lv_date_old_start <= <fs_update>-Validfrom.
          <fs_update>-validfrom = lv_date_old_end + 1.
        ELSE.
          DELETE update_table INDEX sy-tabix.
        ENDIF.
      ENDLOOP.

      IF update_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_weightbsd
               ENTITY ZrSdTWeightbsd UPDATE FIELDS ( Validfrom Validtodate ) WITH update_table.
      ENDIF.
      IF create_table IS NOT INITIAL.
        MODIFY ENTITIES OF zr_sd_t_weightbsd
               ENTITY ZrSdTWeightbsd CREATE AUTO FILL CID WITH create_table
               FAILED failed.
        IF failed IS INITIAL.
          APPEND VALUE #( flag = '1' ) TO et_response.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
