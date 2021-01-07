// <copyright file="HomeController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.VirtualConsult.Controllers
{
    using Microsoft.AspNetCore.Mvc;

    /// <summary>
    /// Controller for home/default views
    /// </summary>
    public class HomeController : Controller
    {
        /// <summary>
        /// The default index view for the Home controller
        /// </summary>
        /// <returns>the view to render</returns>
        public IActionResult Index()
        {
            return this.View();
        }
    }
}