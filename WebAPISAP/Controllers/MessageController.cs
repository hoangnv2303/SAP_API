using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Messaging;
using System.Web.Http.Results;

namespace WebAPISAP.Controllers
{
    public class MessageController : ApiController
    {
        [HttpGet]
        public IHttpActionResult CreateMessage()
        {
            CreateQueue();
            return Ok();
        }

        public static void CreateQueue()
        {
            var emp = new Employee()
            {
                Id = 100,
                Name = "John Doe",
                Hours = DateTime.Now.Millisecond,
                Rate = 21.0
            };
            System.Messaging.Message msg = new System.Messaging.Message();
            msg.Body = emp;
            MessageQueue msgQ = new MessageQueue(".\\Private$\\hoang");
            //MessageQueue msgQ = new MessageQueue("Formatname:Direct=OS:hvlappsweb01-dev\\Private$\\kissQueue");
            msgQ.Send(msg);
        }
        public class Employee
        {
            public int Id;
            public string Name;
            public int Hours;
            public double Rate;
        }
    }
}
