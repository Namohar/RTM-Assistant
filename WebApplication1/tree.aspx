<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="tree.aspx.cs" Inherits="WebApplication1.tree" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript" src="https://www.google.com/jsapi"></script>
    
    <script type="text/javascript">
        google.load("visualization", "1", { packages: ["orgchart"] });
        $("#btnOrgChart").on('click', function (e) {

            $.ajax({
                type: "POST",
                url: "tree.aspx/getOrgData",
                data: '{}',
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: OnSuccess_getOrgData,
                error: OnErrorCall_getOrgData
            });

            function OnSuccess_getOrgData(repo) {

                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Name');
                data.addColumn('string', 'Manager');
                data.addColumn('string', 'ToolTip');

                var response = repo.d;
                for (var i = 0; i < response.length; i++) {
                    var row = new Array();
                    var empName = response[i].Employee;
                    var mgrName = response[i].Manager;
                    var empID = response[i].empID;
                    var mgrID = response[i].mgrID;
                    

                    data.addRows([[{
                        v: empID,
                        f: empName
                    }, mgrID]]);
                }

                var chart = new google.visualization.OrgChart(document.getElementById('chart_div'));
                chart.draw(data, { allowHtml: true });
            }

            function OnErrorCall_getOrgData() {
                console.log("Whoops something went wrong :( ");
            }
            e.preventDefault();
        });

    </script>
</head>
<body>
    <form id="form1" runat="server">
    <input id="btnOrgChart" type="button" value="Click" />
    <div id="chart_div">
      
    </div>
    </form>
</body>
</html>
