using Microsoft.AspNetCore.Mvc.RazorPages;
using FlowLab.Web.Services;

namespace FlowLab.Web.Pages;

public class IndexModel(LaboratoryReadService laboratory) : PageModel
{
    public DashboardSummary Summary { get; private set; } = new(0, 0, 0, 0);

    public void OnGet() => Summary = laboratory.GetDashboardSummary();
}
