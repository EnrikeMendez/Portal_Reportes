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


    <link type="text/css" href="css/logis_style.min.css" rel="stylesheet" />
    <link href="include/logis.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="js/reports.min.js"></script>
    <script language="JavaScript" src="./include/tigra_tables.js"></script>

    <!--ORP: WS AJAX-->
    <script src="js/jquery-1.3.2.min.js"></script>
    <script src="js/main.js"></script>
    <!--ORP: WS AJAX-->

    <style type="text/css">
        .dvHoraActual {
            display: none;
            font-size: small;
            font-weight: bold;
            text-align: right;
            visibility: collapse;
            width: 100%;
        }
    </style>
    <script type="text/javascript">
        var MinutosRecargarPagina = 5;
        function r() { setTimeout(function () { location.reload(); }, 1000); }
        var totLabel = "";
        const fecha = new Date();
        var tot = (fecha.getFullYear() - 2000) + (fecha.getMonth() + 1) + fecha.getDate() + fecha.getHours() + fecha.getMinutes() + fecha.getSeconds();
        tot = tot / 60;
        totLabel = tot.toFixed(1).toString();
    </script>
    <title>Consulta de errores</title>
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
                <li id="imgXls" alt="Exportar" title="Exportar consulta" onclick="GeneraExcel('ConsultaErrores','select_reporte')" class="link_cursor">Exportar consulta
                </li>
            </ul>
            <h2>CONSULTA DE ERRORES
            </h2>
        </div>
    </div>


    <center>
        <table width="98%" border="0" class="tbl-shadow">
            <tr>
                <td>
                    <input id="table-buscar" type="text" class="form-control rounded-txt" placeholder="Escriba algo para filtrar" style="width: 100%;" />
                </td>
            </tr>
        </table>
    </center>
    <br />

    <!--ORP: WS AJAX-->
    <table width="100%" border="0" id="select_reporte" class="tblContent">
        <thead>
            <tr align='center'>
                <th>Nombre</th>
                <th class='width-8p'>ID cron</th>
                <th class='width-8p'>Error</th>
                <!--<th class='width-10p'>Lista Correo</th>-->
                <th>Log</th>
                <th class='width-12p'>Fecha</th>
            </tr>
        </thead>
        <tbody id="print_data">
        </tbody>
    </table>
    <!--ORP: WS AJAX-->


    <div id="clockCounter" name="clockCounter" class="fixedBottomRightLabel">
        La p&aacute;gina se actualizar&aacute en
			<label id="lblTimer" name="lblTimer">60</label>
        <label id="lblMedidaTiempo" name="lblMedidaTiempo">segundo</label>(s).
    </div>


    <script type="text/javascript">
        var reloadTime = MinutosRecargarPagina * 60 * 1000;

        //	< !--
        tigra_tables('select_reporte', 4, 0, '#ffffff', '#ffffcc', '#ffcc66', '#cccccc');
        // -->

        function refresca_pagina() {
            counter();
            showTime();
            setTimeout(function () {
                location.reload();
            }, (reloadTime));
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
        function showTime() {
            var myDate, hours, minutes, seconds, dvHoraActual, dato;
            myDate = new Date();
            hours = myDate.getHours();
            minutes = myDate.getMinutes();
            seconds = myDate.getSeconds();
            if (hours < 10) hours = 0 + hours;
            if (minutes < 10) minutes = "0" + minutes;
            if (seconds < 10) seconds = "0" + seconds;

            dvHoraActual = document.getElementById("HoraActual");
            dato = (hours + ":" + minutes + ":" + seconds);
            dvHoraActual.innerHTML = dato;
            setTimeout("showTime()", 1000);
        }
        function hideLoading() {
            document.getElementById("dvloading").style.display = "none";
            document.getElementById("dvloading").style.visibility = "collapse";
        }

        $TableFilter = function (id, value) {
            var rows = document.querySelectorAll(id + ' tbody tr');

            if (MinutosRecargarPagina != null) { MinutosRecargarPagina = 5; }

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
        function ftn_GetConsultaErrores() {
            const xhr = new XMLHttpRequest();
            const url = urlWebService + "GetConsultaErrores";

            xhr.onreadystatechange = function () {
                if (xhr.readyState == XMLHttpRequest.DONE) {
                    ftn_mostrarResultado(xhr.responseText);
                }
            }

            xhr.open("GET", url, true);
            xhr.send();
        }



        function ftn_mostrarResultado(wsResponseText) {
            var objResult = JSON.parse(wsResponseText);
            var info = objResult.GetConsultaErroresResult;
            var arrayData = JSON.parse(info);

            var i = 0;
            var htmlTable = "";


            if (arrayData.length == 0) {

                htmlTable = htmlTable + "<tr class='center'>";
                htmlTable = htmlTable + "	<td colspan='5' class='center'>";
                htmlTable = htmlTable + "		No hay Errores registrados.";
                htmlTable = htmlTable + "	</td>";
                htmlTable = htmlTable + "</tr>";
            }
            else {

                for (i = 0; i < arrayData.length; i++) {

                    htmlTable = "";

                    htmlTable = htmlTable + "<tr> \n\n";

                   // console.log(arrayData[0, i].toString());

                    htmlTable = htmlTable + "<td>";
                    if (arrayData[i].NAME.includes("@", 1) == false) {
                        htmlTable = htmlTable + arrayData[i].NAME + "\n\n";
                    } else {
                        htmlTable = htmlTable + "<i>Reporte generado bajo demanda</i> \n\n";
                    }
                    htmlTable = htmlTable + "</td>";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + "<font class='carmin'>";
                    htmlTable = htmlTable + arrayData[i].ID_CRON;
                    htmlTable = htmlTable + "</font>";
                    htmlTable = htmlTable + "</td>";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].ID_CHRON_ERROR;
                    htmlTable = htmlTable + "</td>";

                    htmlTable = htmlTable + "<!--<td>";
                    htmlTable = htmlTable + arrayData[i].TIPO_ERROR;
                    htmlTable = htmlTable + "</td>-->";

                    htmlTable = htmlTable + "<td>";
                    htmlTable = htmlTable + arrayData[i].LOG;
                    htmlTable = htmlTable + "</td>";

                    htmlTable = htmlTable + "<td align='center'>";
                    htmlTable = htmlTable + arrayData[i].HORA;
                    htmlTable = htmlTable + "</td>";

                    htmlTable = htmlTable + "</tr>";

                    $("#print_data").append(htmlTable);
                }
            }
        }
        $(document).ready(ftn_GetConsultaErrores);
            //< !--ORP: WS AJAX-- >
    </script>

    <div id="HoraActual" name="HoraActual" class="dvHoraActual"></div>
</body>
</html>
