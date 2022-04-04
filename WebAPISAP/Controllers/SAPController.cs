using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WebAPISAP.Common;

namespace WebAPISAP.Controllers
{
    public class SAPController : ApiController
    {
        public SAPHelper _sap;
        [HttpPost]
        [Route("api/DS/get")]
        public IHttpActionResult GetDeliverySales([FromBody]dynamic value)
        {
            _sap = new SAPHelper();
            WriteLogs.Write("=== Delivery Sale SAP ====", "Start");
            DataTable dsTable = _sap.DownloadDeliverySale(value.plant.Value, value.from.Value, value.to.Value);
            WriteLogs.Write("=== Delivery Sale SAP ====", "Done");
            WriteLogs.Write("=== Total Record: ", dsTable.Rows.Count.ToString());
            if (dsTable != null && dsTable.Rows.Count > 0)
            {
                return Ok(dsTable);
            }
            return Ok();
        }

        [HttpPost]
        [Route("api/Boom/get")]
        public IHttpActionResult GetBoom([FromBody]dynamic value)
        {
            _sap = new SAPHelper();
            DataTable dsTable = _sap.DownloadBoom(value.plant.Value, value.parentMaterial.Value);
            if (dsTable != null && dsTable.Rows.Count > 0)
            {
                return Ok(dsTable);
            }
            WriteLogs.Write("=== Download Boom SAP Error====", $"=== Plant: {value.plant.Value}, Parent Material: {value.parentMaterial.Value}");
            return Ok();
        }

        [HttpPost]
        [Route("api/ListTO/get")]
        public IHttpActionResult GetListTo([FromBody]dynamic value)
        {
            _sap = new SAPHelper();
            WriteLogs.Write("=== Get List TO SAP ====", "Start");
            DataTable dsTable = _sap.GetListTO(value.fromDate.Value, value.toDate.Value);
            WriteLogs.Write("=== Get List TO SAP ====", "Done");
            WriteLogs.Write("=== Total Record: ", dsTable.Rows.Count.ToString());
            if (dsTable != null && dsTable.Rows.Count > 0)
            {
                return Ok(dsTable);
            }
            return Ok();
        }
    }
}

