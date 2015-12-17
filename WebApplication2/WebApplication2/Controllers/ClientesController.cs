using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApplication2.Models;

namespace WebApplication2.Controllers
{
    public class ClientesController : Controller
    {
        Clientes cli;
        // GET: Clientes
        public ActionResult Index()
        {
            cli = new Clientes();

            cli.id = 1;
            cli.Checa = true;

          Clientes cli1 = new Clientes();
            cli1.id = 2;
            cli1.Checa = false;

            List<Clientes> lista = new List<Clientes>();

            lista.Add(cli);
            lista.Add(cli1);

            return View(lista);
        }

        [HttpGet]
        [AcceptVerbs(HttpVerbs.Get)]
        public ActionResult Alterar(IEnumerable<Clientes> clientes) {

            return View();
        }
    }
}