export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return Response.json({ ok: true });
    }

    if (url.pathname === "/api/reports" && request.method === "POST") {
      try {
        const report = await request.json();

        const required = [
          "submission_date",
          "report_number",
          "person_name",
          "occurrence_location",
          "severity",
          "cause_analysis",
          "recurrence_prevention"
        ];

        for (const field of required) {
          if (!report[field]) {
            return Response.json(
              { success: false, error: `${field} is required` },
              { status: 400 }
            );
          }
        }

        const result = await env.DB.prepare(
          `INSERT INTO accident_reports
           (submission_date, report_number, person_name, occurrence_location, severity, cause_analysis, recurrence_prevention)
           VALUES (?, ?, ?, ?, ?, ?, ?)`
        )
          .bind(
            report.submission_date,
            report.report_number,
            report.person_name,
            report.occurrence_location,
            report.severity,
            report.cause_analysis,
            report.recurrence_prevention
          )
          .run();

        return Response.json({ success: true, id: result.meta.last_row_id });
      } catch (error) {
        return Response.json(
          { success: false, error: error.message },
          { status: 500 }
        );
      }
    }

    if (url.pathname === "/api/reports" && request.method === "GET") {
      const { results } = await env.DB.prepare(
        `SELECT * FROM accident_reports ORDER BY id DESC`
      ).all();
      return Response.json({ success: true, reports: results });
    }

    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Not Found", { status: 404 });
  }
};
