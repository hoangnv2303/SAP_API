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
        [Route("api/DS/get")]
        public IHttpActionResult GetDeliverySales()
        {
            _sap = new SAPHelper();
            var now = DateTime.Now.AddDays(-20);
            DataTable dsTable = _sap.DownloadDeliverySale("FI20", now, now.AddDays(1));
            if (dsTable != null && dsTable.Rows.Count > 0)
            {
                return Ok(dsTable);
            }
            return BadRequest();
        }
    }
}
