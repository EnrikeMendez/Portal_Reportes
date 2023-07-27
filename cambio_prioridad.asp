<%@  language="VBScript" %>
<% option explicit %>
<!--#include file="include/include.asp"-->

<%
'admin of logis web site :
'modificacion de reportes
	Response.Expires = 0
	call check_session()

%>
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html;" charset="iso-8859-1" />

    <!--ORP: WS AJAX-->
    <!--call print_style()-->
    <link href="css/print_style.css" type="text/css" rel="stylesheet" />
    <!--ORP: WS AJAX-->


    <link href="include/logis.css" type="text/css" rel="stylesheet" />
    <link href="css/logis_style.css" type="text/css" rel="stylesheet" />
    <script language="JavaScript" src="./include/tigra_tables.js"></script>
    <script type="text/javascript" src="js/reports.min.js"></script>
    <!--ORP: WS AJAX-->
    <script src="js/jquery-1.3.2.min.js"></script>
    <script src="js/main.js"></script>
    <!--ORP: WS AJAX-->

    <script type="text/javascript">



        var MinutosRecargarPagina = 5;

        var totLabel = "";
        const fecha = new Date();
        var tot = (fecha.getFullYear() - 2000) + (fecha.getMonth() + 1) + fecha.getDate() + fecha.getHours() + fecha.getMinutes() + fecha.getSeconds();
        tot = tot / 60;
        totLabel = tot.toFixed(1).toString();
    </script>
    <title>Cambio de prioridad</title>
</head>
<body onload="refresca_pagina();">
    <div id="dvloading" style="display: block!important; visibility: visible!important;">
        <center>Procesando </center>
        <center>
            <img alt=". . ." id="imgPuntos" src="images/puntosSuspensivos.gif" /></center>
    </div>
    <div class="contenedorMenu">
        <div class="dvMenu">
            <ul id="menu">
                <div class="logo-logis">
                    <img src="images/logo-logis-s.png" style="height: 50px;" />
                </div>
                <li onclick="window.location.href='menu.asp';" class="link_cursor">Inicio
                </li>
                <li id="imgXls" alt="Exportar" title="Exportar consulta" onclick="GeneraExcel('CambioPrioridad','select_reporte')" class="link_cursor">Exportar consulta
                </li>
            </ul>
            <h2>CAMBIO DE PRIORIDAD
            </h2>
        </div>
    </div>



    <form action="" method="">
        <center>
            <table width="98%" border="0" class="tbl-shadow">
                <tr>
                    <td colspan="2">
                        <input id="table-buscar" type="text" class="form-control rounded-txt" placeholder="Escriba algo para filtrar" style="width: 100%;" />
                    </td>
                </tr>
                <tr>
                    <td class="width-15p">
                        <label>Prioridad</label>
                        <select name="select_prioridad" id="select_prioridad" class="form-control rounded-cmb">
                            <option value="0" selected>0</option>
                            <option value="1">1</option>
                            <option value="2">2</option>
                            <option value="3">3</option>
                            <option value="4">4</option>
                            <option value="5">5</option>
                            <option value="6">6</option>
                            <option value="7">7</option>
                            <option value="8">8</option>
                            <option value="9">9</option>
                        </select>
                    </td>
                    <td>

                        <!--ORP: WS AJAX-->
                        <!--<button type="submit" class="rounded-btn">Guardar</button>-->
                        <button id="btn_guardar" class="rounded-btn" onclick="ftn_GettModificaCambioPrioridad()">Guardar</button>
                        <!--ORP: WS AJAX-->

                    </td>
                </tr>
            </table>
        </center>
        <br />
        <table width="100%" border="0" id="select_reporte" class="tblContent">
            <thead>
                <tr align="center">
                    <th>&nbsp;</th>
                    <th>ID rep</th>
                    <th>Nombre</th>
                    <th>Prioridad</th>
                    <th>Cliente</th>
                    <th>ID cron</th>
                    <th>Nombre detalle</th>
                    <th>Fecha</th>
                    <th>Email</th>
                </tr>
            </thead>
            <tbody id="print_data">
            </tbody>
        </table>
    </form>

    <div id="clockCounter" name="clockCounter" class="fixedBottomRightLabel">
        La p&aacute;gina se actualizar&aacute en
			<label id="lblTimer" name="lblTimer">60</label>
        <label id="lblMedidaTiempo" name="lblMedidaTiempo">segundo</label>(s).
    </div>

    <script type="text/javascript">
        var reloadTime = MinutosRecargarPagina * 60 * 1000;

        //<!--
        tigra_tables('select_reporte', 4, 0, '#ffffff', '#ffffcc', '#ffcc66', '#cccccc');
        // -->

        function refresca_pagina() {
            counter();
            setTimeout(function () {
                location.reload();
            }, reloadTime);
            hideLoading();
        }
        function counter() {
            var dNow = new Date();
            var redColorSeconds = 5;
            var t = Math.round(reloadTime / 1000);
            var lblTimer = document.getElementById("lblTimer");
            var lblText = document.getElementById("clockCounter");
            var lblMedidaTiempo = document.getElementById("lblMedidaTiempo");

            try { redColorSeconds = ((dNow.getDate() + dNow.getMonth()) * 0.7) + 1; }
            catch { redColorSeconds = redColorSeconds + 1; }

            lblTimer.innerHTML = "<b>" + MinutosRecargarPagina + "</b>";
            lblMedidaTiempo.innerHTML = "minuto";

            window.setInterval(function () {
                lblTimer.innerHTML = t - 1;
                t--;

                if (t <= 60) {
                    lblMedidaTiempo.innerHTML = "segundo";
                }
                if (t < redColorSeconds) {
                    lblTimer.style.color = "red";
                }
                if (t <= (redColorSeconds * 1.5)) {
                    lblText.style.display = "block";
                }
                else {
                    lblText.style.display = "none";
                }
            }, 1000);
        }
        function showLoading() {
            document.getElementById("dvloading").style.display = "";
            document.getElementById("dvloading").style.visibility = "visible";
        }
        function hideLoading() {
            document.getElementById("dvloading").style.display = "none";
            document.getElementById("dvloading").style.visibility = "collapse";
        }

        $TableFilter = function (id, value) {
            var rows = document.querySelectorAll(id + ' tbody tr');

            for (var i = 0; i < rows.length; i++) {
                var showRow = false;

                var row = rows[i];
                row.style.display = 'none';

                for (var x = 0; x < row.childElementCount; x++) {
                    if (row.children[x].textContent.toLowerCase().indexOf(value.toLowerCase().trim()) > -1) {
                        showRow = true;
                        break;
                    }
                }

                if (showRow) {
                    row.style.display = null;
                }
            }
        }

        document.querySelector("#table-buscar").onkeyup = function () {
            $TableFilter("#select_reporte", this.value);
        }


        //< !--ORP: WS AJAX-- >
        function ftn_GetConsultaCambioPrioridad() {
            const xhr = new XMLHttpRequest();
            //const url = "http://localhost:50899/Report_Service.svc/GetConsultaCambioPrioridad";
            const url = urlWebService + "GetConsultaCambioPrioridad";

            xhr.onreadystatechange = function () {
                if (xhr.readyState == XMLHttpRequest.DONE) {
                    ftn_consulta_cambio_prioridad(xhr.responseText);
                }
            }

            xhr.open("GET", url, true);
            xhr.send();
        }

        function ftn_GettModificaCambioPrioridad() {
            const xhr = new XMLHttpRequest();

            var id_crons = "";
            var prioridad = 0;

            prioridad = document.getElementById("select_prioridad").value;
            //id_crons = document.getElementById("reporte[]").value;
            id_crons = document.getElementById("arr").value;

            if (id_crons != "") {
                //const url = "http://localhost:51687/Report_Service.svc/GetModificaCambioPrioridad?id_crons = "+ id_crons +" ";
                const url = urlWebService + "GetModificaCambioPrioridad?id_crons=" + id_crons + "&prioridad= " + prioridad + "";

                xhr.onreadystatechange = function () {
                    if (xhr.readyState == XMLHttpRequest.DONE) {
                        ftn_modifica_cambio_prioridad(xhr.responseText);
                    }
                }
                xhr.open("GET", url, true);
                xhr.send();
            }
            else {
                alert("Seleccione un reporte primero.");
            }

            
        }



        function ftn_consulta_cambio_prioridad(wsResponseText) {

            var objResult = JSON.parse(wsResponseText);
            var info = objResult.GetConsultaCambioPrioridadResult;
            var arrayData = JSON.parse(info);

            var i = 0;
            var htmlTable = "";

            if (arrayData.length == 0) {
                htmlTable = htmlTable + "<tr class='center' >";
                htmlTable = htmlTable + "<td colspan='9' class='center'>";
                htmlTable = htmlTable + " No hay Reportes en ejecuci&oacute;n.";
                htmlTable = htmlTable + "</td>";
                htmlTable = htmlTable + "</tr>";

                $("#print_data").append(htmlTable);

            } else {

                for (i = 0; i < arrayData.length; i++) {

                    //htmlTable = "";

                    htmlTable = htmlTable + "<tr>";

                    htmlTable = htmlTable + " <td align='center'><input type='checkbox' name='reporte[]' id='reporte[]' value='" + arrayData[i].ID_CRON + "'></td> \n\n"


                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].ID_REP;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td> <font class='carmin'>";
                    htmlTable = htmlTable + arrayData[i].NOMBRE_REPORTE;
                    htmlTable = htmlTable + "</font> </td> \n\n";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].PRIORITE;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].CLIENTE;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].ID_CRON;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td>";
                    htmlTable = htmlTable + arrayData[i].NOMBRE_DETALLE;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td>";
                    htmlTable = htmlTable + arrayData[i].HORA_CREACION;
                    htmlTable = htmlTable + "</td> \n\n";

                    htmlTable = htmlTable + "<td>";
                    if (arrayData[i].DEST_MAIL != "") {
                        htmlTable = htmlTable + arrayData[i].DEST_MAIL;
                    }
                    htmlTable = htmlTable + "</td> \n\n";


                    htmlTable = htmlTable + "</tr> \n\n";
                }

                $("#print_data").append(htmlTable);
            }

        }
        $(document).ready(ftn_GetConsultaCambioPrioridad);

        function ftn_modifica_cambio_prioridad(wsResponseText) {
            var objResult = JSON.parse(wsResponseText);
            var info = objResult.GetConsultaCambioPrioridadResult;
            alert(info);
        }
            //< !--ORP: WS AJAX-- >



        //function select_check_reporte() {
        //    $('[name="reporte[]"]').click(function () {
        //        var arr = $('[name="reporte[]"]:checked').map(function () {
        //            return this.value;
        //        }).get();
        //        var str = arr.join(',');
        //        $('#arr').text(JSON.stringify(arr));
        //        $('#str').text(str);
        //        <input type="hidden" name="arr" id="arr" value="arr" />
        //    });
        //}



    </script>
</body>
</html>
