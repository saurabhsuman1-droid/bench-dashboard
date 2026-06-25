import { useState, useEffect, useRef, useCallback } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import Chart from "chart.js/auto";

const CLIENT_ID   = import.meta.env.VITE_AAD_CLIENT_ID;
const USE_CLOUD   = !!CLIENT_ID;
const GRAPH_FILE  = "https://graph.microsoft.com/v1.0/me/drive/root:/BenchDashboard/bench-data.json:/content";
const TOKEN_SCOPE = ["Files.ReadWrite"];

export default function App() {
  const migrate = (data) =>
    data.map((a, i) => ({ id: a.id ?? i + 1, archived: a.archived ?? false, ...a }));

  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const authBusy = inProgress !== InteractionStatus.None;

  const getToken = useCallback(async () => {
    try {
      const res = await instance.acquireTokenSilent({ scopes: TOKEN_SCOPE, account: accounts[0] });
      return res.accessToken;
    } catch {
      const res = await instance.acquireTokenPopup({ scopes: TOKEN_SCOPE });
      return res.accessToken;
    }
  }, [instance, accounts]);

  const [associates, setAssociates] = useState(() => {
    if (USE_CLOUD) return [];
    const saved = localStorage.getItem("benchData");
    return migrate(saved ? JSON.parse(saved) : []);
  });

  const [coursePool, setCoursePool] = useState(() => {
    if (USE_CLOUD) return ["GenAI", "Claude"];
    const saved = localStorage.getItem("coursePool");
    return saved ? JSON.parse(saved) : ["GenAI", "Claude"];
  });

  const [loading,   setLoading]   = useState(USE_CLOUD);
  const [syncing,   setSyncing]   = useState(false);
  const [syncError, setSyncError] = useState(false);

  const [isAdmin, setIsAdmin] = useState(false);
  const [managerView, setManagerView] = useState(true);
  const [showAnalytics, setShowAnalytics] = useState(false);
  const [expandedId, setExpandedId] = useState(null);
  const [name, setName] = useState("");
  const [joinDate, setJoinDate] = useState("");
  const [newCourse, setNewCourse] = useState("");

  const chart1Ref  = useRef(null);
  const chart2Ref  = useRef(null);
  const chartInst1 = useRef(null);
  const chartInst2 = useRef(null);
  const saveTimer   = useRef(null);
  const skipSave    = useRef(false);
  const cloudReady  = useRef(!USE_CLOUD); // flips true once initial fetch completes

  // ── Load from OneDrive on auth (cloud mode only) ──
  useEffect(() => {
    if (!USE_CLOUD || !isAuthenticated) return;
    getToken()
      .then(token => fetch(GRAPH_FILE, { headers: { Authorization: `Bearer ${token}` } }))
      .then(r => r.status === 404 ? null : r.json())
      .then(record => {
        if (record?.associates) setAssociates(migrate(record.associates));
        if (record?.coursePool) setCoursePool(record.coursePool);
        if (record) skipSave.current = true; // only skip echo when data was actually loaded
        cloudReady.current = true;
      })
      .catch(() => setSyncError(true))
      .finally(() => setLoading(false));
  }, [isAuthenticated, getToken]);

  // ── Persist on every data change ──
  useEffect(() => {
    if (!USE_CLOUD) {
      localStorage.setItem("benchData",  JSON.stringify(associates));
      localStorage.setItem("coursePool", JSON.stringify(coursePool));
      return;
    }
    if (!cloudReady.current) return;
    if (skipSave.current) { skipSave.current = false; return; }

    setSyncing(true);
    setSyncError(false);
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        const token = await getToken();
        await fetch(GRAPH_FILE, {
          method: "PUT",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ associates, coursePool }),
        });
        setSyncing(false);
      } catch {
        setSyncing(false);
        setSyncError(true);
      }
    }, 1200);
  }, [associates, coursePool, getToken]);

  // Render analytics charts whenever panel opens or data changes
  useEffect(() => {
    if (!showAnalytics) return;

    const act = associates.filter(a => !a.archived);
    const allClients = act.flatMap(a => a.clients);

    if (chartInst1.current) { chartInst1.current.destroy(); chartInst1.current = null; }
    if (chartInst2.current) { chartInst2.current.destroy(); chartInst2.current = null; }

    if (chart1Ref.current) {
      const dist = coursePool.map(c => ({
        course: c,
        notStarted: act.filter(a => (a.courses[c] || 0) === 0).length,
        inProgress: act.filter(a => (a.courses[c] || 0) > 0 && (a.courses[c] || 0) < 80).length,
        completed: act.filter(a => (a.courses[c] || 0) >= 80).length,
      }));
      chartInst1.current = new Chart(chart1Ref.current, {
        type: "bar",
        data: {
          labels: dist.map(d => d.course),
          datasets: [
            { label: "Not Started", data: dist.map(d => d.notStarted), backgroundColor: "#fca5a5", borderRadius: 3 },
            { label: "In Progress (1–79%)", data: dist.map(d => d.inProgress), backgroundColor: "#fde68a", borderRadius: 3 },
            { label: "Completed (≥80%)", data: dist.map(d => d.completed), backgroundColor: "#86efac", borderRadius: 3 },
          ],
        },
        options: {
          plugins: {
            title: { display: true, text: "Course Completion Distribution", font: { size: 13 }, color: "#1e1b4b" },
            legend: { position: "top", labels: { font: { size: 11 }, boxWidth: 12, padding: 10 } },
          },
          scales: {
            x: { stacked: true },
            y: { stacked: true, ticks: { stepSize: 1 }, min: 0, title: { display: true, text: "# Associates" } },
          },
        },
      });
    }

    if (chart2Ref.current) {
      chartInst2.current = new Chart(chart2Ref.current, {
        type: "bar",
        data: {
          labels: ["Submitted", "Internal Done", "Client Round", "Selected"],
          datasets: [{
            data: [
              allClients.length,
              allClients.filter(c => c.internal === "Done").length,
              allClients.filter(c => c.client === "Done").length,
              allClients.filter(c => c.outcome === "Selected").length,
            ],
            backgroundColor: ["#818cf8", "#38bdf8", "#fb923c", "#16a34a"],
            borderRadius: 5,
          }],
        },
        options: {
          plugins: {
            legend: { display: false },
            title: { display: true, text: "Client Interview Funnel (Team)", font: { size: 13 }, color: "#1e1b4b" },
          },
          scales: { y: { ticks: { stepSize: 1 }, min: 0 } },
        },
      });
    }

    return () => {
      if (chartInst1.current) { chartInst1.current.destroy(); chartInst1.current = null; }
      if (chartInst2.current) { chartInst2.current.destroy(); chartInst2.current = null; }
    };
  }, [showAnalytics, associates, coursePool]);

  // ── derived ──
  const active   = associates.filter(a => !a.archived);
  const archived = associates.filter(a => a.archived);

  const days = (d) => Math.floor((new Date() - new Date(d)) / 86400000);

  const avgCompletion = (a) => {
    const vals = Object.values(a.courses);
    return Math.round(vals.reduce((s, x) => s + x, 0) / (vals.length || 1));
  };

  const courseChip = (v) =>
    v >= 80 ? "bg-green-100 text-green-700" :
    v >= 40 ? "bg-yellow-100 text-yellow-700" :
    v >  0  ? "bg-orange-100 text-orange-700" :
              "bg-gray-100 text-gray-400";

  const getStats = () => {
    const allClients = active.flatMap(a => a.clients);
    const byProject = {};
    active.forEach(a => a.clients.forEach(c => {
      if (c.projectCode) byProject[c.projectCode] = (byProject[c.projectCode] || 0) + 1;
    }));
    return {
      total:            active.length,
      proposed:         active.filter(a => a.clients.length > 0).length,
      inPipeline:       active.filter(a => a.clients.some(c => c.internal === "Done" || c.client === "Done")).length,
      placed:           active.filter(a => a.clients.some(c => c.outcome === "Selected")).length,
      totalSubmissions: allClients.length,
      selected:         allClients.filter(c => c.outcome === "Selected").length,
      rejected:         allClients.filter(c => c.outcome === "Rejected").length,
      avgDays:          active.length ? Math.round(active.reduce((s, a) => s + days(a.joinDate), 0) / active.length) : 0,
      selectionRate:    allClients.length ? Math.round(allClients.filter(c => c.outcome === "Selected").length / allClients.length * 100) : 0,
      byProject,
    };
  };

  // ── mutations (id-based, immutable) ──
  const nextId = () => associates.length ? Math.max(...associates.map(a => a.id || 0)) + 1 : 1;

  const syncCourses = (updated) =>
    setAssociates(prev => prev.map(a => {
      const obj = {};
      updated.forEach(c => obj[c] = a.courses[c] || 0);
      return { ...a, courses: obj };
    }));

  const addAssociate = () => {
    if (!name || !joinDate) return;
    const obj = {};
    coursePool.forEach(c => obj[c] = 0);
    setAssociates([...associates, { id: nextId(), name, joinDate, courses: obj, clients: [], archived: false }]);
    setName(""); setJoinDate("");
  };

  const updateCourse = (id, c, v) =>
    setAssociates(prev => prev.map(a => a.id === id ? { ...a, courses: { ...a.courses, [c]: parseInt(v) } } : a));

  const addClient = (id) =>
    setAssociates(prev => prev.map(a => a.id !== id ? a : {
      ...a,
      clients: [...a.clients, { name: "", projectCode: "", internal: "Not Started", client: "Not Started", outcome: "Pending" }],
    }));

  const removeClient = (id, j) =>
    setAssociates(prev => prev.map(a => a.id !== id ? a : { ...a, clients: a.clients.filter((_, i) => i !== j) }));

  const updateClient = (id, j, f, v) =>
    setAssociates(prev => prev.map(a => {
      if (a.id !== id) return a;
      return { ...a, clients: a.clients.map((c, i) => i === j ? { ...c, [f]: v } : c) };
    }));

  const addCourse = () => {
    if (!newCourse || coursePool.includes(newCourse)) return;
    const updated = [...coursePool, newCourse];
    setCoursePool(updated);
    syncCourses(updated);
    setNewCourse("");
  };

  const removeCourse = (c) => {
    const updated = coursePool.filter(x => x !== c);
    setCoursePool(updated);
    syncCourses(updated);
  };

  const archiveAssociate = (id) => {
    setAssociates(prev => prev.map(a => a.id === id ? { ...a, archived: true } : a));
    setExpandedId(null);
  };

  const unarchiveAssociate = (id) =>
    setAssociates(prev => prev.map(a => a.id === id ? { ...a, archived: false } : a));

  // ── PDF export ──
  const renderChart = (config, width = 1100, height = 540) => new Promise(res => {
    const c = document.createElement("canvas");
    c.width = width; c.height = height;
    c.style.cssText = "position:fixed;left:-9999px";
    document.body.appendChild(c);
    const chart = new Chart(c, config);
    setTimeout(() => { const img = c.toDataURL("image/png"); chart.destroy(); document.body.removeChild(c); res(img); }, 650);
  });

  const downloadPDF = async () => {
    const doc  = new jsPDF();
    const stats = getStats();
    const allClients = active.flatMap(a => a.clients);

    // ── Page 1: KPI summary + table ──
    doc.setFontSize(18); doc.setTextColor(79, 70, 229);
    doc.text("Bench Dashboard Report", 14, 16);
    doc.setFontSize(9); doc.setTextColor(120);
    doc.text(`Generated: ${new Date().toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })}  |  Active: ${stats.total}  |  Archived: ${archived.length}`, 14, 23);

    const kpis = [
      { label: "TOTAL BENCH",  value: stats.total,       fill: [239,246,255], text: [30,64,175]  },
      { label: "PROPOSED",     value: stats.proposed,    fill: [240,249,255], text: [3,105,161]  },
      { label: "IN PIPELINE",  value: stats.inPipeline,  fill: [254,249,195], text: [133,77,14]  },
      { label: "PLACED",       value: stats.placed,      fill: [220,252,231], text: [21,128,61]  },
    ];
    kpis.forEach((b, i) => {
      const x = 14 + i * 47;
      doc.setFillColor(...b.fill);
      doc.roundedRect(x, 28, 44, 18, 2, 2, "F");
      doc.setFontSize(14); doc.setTextColor(...b.text);
      doc.text(String(b.value), x + 22, 38, { align: "center" });
      doc.setFontSize(7); doc.text(b.label, x + 22, 43, { align: "center" });
    });

    doc.setTextColor(0);
    autoTable(doc, {
      head: [["Name", "Days", "Proposals", "Outcome", "Avg %", "Course Progress"]],
      body: active.map(a => {
        const sel = a.clients.filter(c => c.outcome === "Selected").length;
        const rej = a.clients.filter(c => c.outcome === "Rejected").length;
        return [
          a.name,
          days(a.joinDate),
          a.clients.length,
          sel > 0 ? `✓ ${sel} Selected` : rej > 0 ? `✗ ${rej} Rejected` : a.clients.length > 0 ? "Pending" : "—",
          `${avgCompletion(a)}%`,
          Object.entries(a.courses).map(([k, v]) => `${k}: ${v}%`).join("  |  "),
        ];
      }),
      startY: 52,
      styles: { fontSize: 8, cellPadding: 3 },
      headStyles: { fillColor: [79, 70, 229], fontSize: 8 },
      columnStyles: { 2: { halign: "center" }, 4: { halign: "center" } },
    });

    // ── Page 2: 2 team-level charts (scale to any team size) ──
    doc.addPage();
    doc.setFontSize(14); doc.setTextColor(79, 70, 229);
    doc.text("Visual Summary", 14, 13);

    // Chart 1 — Course completion distribution (stacked bar per course)
    const dist = coursePool.map(c => ({
      course: c,
      notStarted: active.filter(a => (a.courses[c] || 0) === 0).length,
      inProgress: active.filter(a => (a.courses[c] || 0) > 0 && (a.courses[c] || 0) < 80).length,
      completed:  active.filter(a => (a.courses[c] || 0) >= 80).length,
    }));

    const img1 = await renderChart({
      type: "bar",
      data: {
        labels: dist.map(d => d.course),
        datasets: [
          { label: "Not Started (0%)",   data: dist.map(d => d.notStarted), backgroundColor: "#fca5a5", borderRadius: 3 },
          { label: "In Progress (1–79%)", data: dist.map(d => d.inProgress), backgroundColor: "#fde68a", borderRadius: 3 },
          { label: "Completed (≥80%)",    data: dist.map(d => d.completed),  backgroundColor: "#86efac", borderRadius: 3 },
        ],
      },
      options: {
        responsive: false,
        plugins: {
          legend: { position: "top", labels: { font: { size: 14 }, padding: 16, boxWidth: 14 } },
          title: { display: true, text: "Course Completion — How Many Associates Per Stage", font: { size: 17 }, color: "#1e1b4b", padding: { bottom: 14 } },
        },
        scales: {
          x: { stacked: true, ticks: { font: { size: 14 } } },
          y: { stacked: true, ticks: { stepSize: 1, font: { size: 13 } }, min: 0, title: { display: true, text: "# Associates", font: { size: 13 } } },
        },
      },
    });

    // Chart 2 — Team pipeline funnel (aggregate, not per-person)
    const img2 = await renderChart({
      type: "bar",
      data: {
        labels: ["Submitted", "Internal Done", "Client Round Done", "Selected"],
        datasets: [{
          data: [
            allClients.length,
            allClients.filter(c => c.internal === "Done").length,
            allClients.filter(c => c.client === "Done").length,
            allClients.filter(c => c.outcome === "Selected").length,
          ],
          backgroundColor: ["#818cf8", "#38bdf8", "#fb923c", "#16a34a"],
          borderRadius: 6,
        }],
      },
      options: {
        responsive: false,
        plugins: {
          legend: { display: false },
          title: {
            display: true,
            text: `Client Interview Funnel  •  Selection Rate: ${stats.selectionRate}%  •  ${stats.selected} of ${stats.totalSubmissions} placed`,
            font: { size: 16 }, color: "#1e1b4b", padding: { bottom: 14 },
          },
        },
        scales: {
          x: { ticks: { font: { size: 14 } } },
          y: { ticks: { stepSize: 1, font: { size: 13 } }, min: 0 },
        },
      },
    });

    doc.addImage(img1, "PNG", 10, 18,  190, 118);
    doc.addImage(img2, "PNG", 10, 142, 190, 100);

    doc.save("bench_report.pdf");
  };

  const stats = getStats();

  const headerBtn = (on) =>
    `px-3 py-1.5 rounded text-sm font-medium border transition-all ${
      on ? "bg-white text-indigo-700 border-white shadow font-semibold"
         : "text-white border-white/30 hover:bg-white/10 hover:border-white/60"
    }`;

  const switchMode = (fn) => { fn(); setShowAnalytics(false); setExpandedId(null); };

  if (USE_CLOUD && !isAuthenticated) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center">
      <div className="text-center">
        <div className="w-12 h-12 bg-indigo-600 rounded-xl flex items-center justify-center mx-auto mb-4">
          <span className="text-white text-xl font-bold">B</span>
        </div>
        <h1 className="text-2xl font-bold text-gray-800 mb-1">Bench Dashboard</h1>
        <p className="text-gray-500 text-sm mb-6">Sign in with your EPAM account to continue</p>
        <button
          disabled={authBusy}
          onClick={() => instance.loginRedirect({ scopes: TOKEN_SCOPE })}
          className="px-6 py-2.5 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium text-sm disabled:opacity-50 disabled:cursor-not-allowed"
        >
          {authBusy ? "Signing in…" : "Sign in with Microsoft"}
        </button>
      </div>
    </div>
  );

  if (loading) return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center">
      <div className="text-center">
        <div className="w-8 h-8 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mx-auto mb-3" />
        <p className="text-gray-500 text-sm">Loading dashboard…</p>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-gray-50">

      {/* Header */}
      <div className="bg-gradient-to-r from-indigo-600 to-purple-600 text-white px-6 py-3 flex justify-between items-center shadow-md">
        <div>
          <h1 className="text-lg font-bold tracking-tight">Bench Dashboard</h1>
          <p className="text-xs text-indigo-200">
            Associate Tracking & Pipeline
            {USE_CLOUD && (
              <span className={`ml-2 ${syncError ? "text-red-300" : syncing ? "text-indigo-300" : "text-green-300"}`}>
                {syncError ? "· ⚠ sync failed" : syncing ? "· saving…" : "· synced"}
              </span>
            )}
          </p>
        </div>
        <div className="flex gap-2 items-center">
          <button onClick={downloadPDF} className="bg-green-500 hover:bg-green-400 text-white px-3 py-1.5 rounded text-sm font-medium border border-green-400 transition-colors shadow-sm">
            Export PDF
          </button>
          {!isAdmin && (
            <button onClick={() => setShowAnalytics(v => !v)} className={headerBtn(showAnalytics)}>
              Analytics
            </button>
          )}
          {!isAdmin && (
            <button onClick={() => switchMode(() => setManagerView(v => !v))} className={headerBtn(!managerView)}>
              {managerView ? "Edit" : "View"}
            </button>
          )}
          <button onClick={() => switchMode(() => setIsAdmin(v => !v))} className={headerBtn(isAdmin)}>
            Admin
          </button>
        </div>
      </div>

      <div className="p-6 space-y-4 max-w-5xl mx-auto">

        {/* KPI Cards — always visible in non-admin mode */}
        {!isAdmin && (
          <div className="grid grid-cols-4 gap-3">
            {[
              { label: "TOTAL BENCH",  value: stats.total,      sub: "active associates",    cls: "bg-white border-gray-100 text-gray-800 text-gray-500" },
              { label: "PROPOSED",     value: stats.proposed,   sub: "with client(s) added", cls: "bg-blue-50 border-blue-200 text-blue-700 text-blue-500" },
              { label: "IN PIPELINE",  value: stats.inPipeline, sub: "rounds underway",      cls: "bg-yellow-50 border-yellow-200 text-yellow-700 text-yellow-600" },
              { label: "PLACED",       value: stats.placed,     sub: "selected by client",   cls: "bg-green-50 border-green-200 text-green-700 text-green-500" },
            ].map(k => {
              const [c0, c1, c2, c3] = k.cls.split(" ");
              return (
                <div key={k.label} className={`p-4 rounded-lg border shadow-sm text-center ${c0} ${c1}`}>
                  <div className={`text-3xl font-bold ${c2}`}>{k.value}</div>
                  <div className={`text-xs font-semibold mt-1 tracking-wide ${c2}`}>{k.label}</div>
                  <div className={`text-xs mt-0.5 opacity-70 ${c3}`}>{k.sub}</div>
                </div>
              );
            })}
          </div>
        )}

        {/* Analytics Panel */}
        {showAnalytics && !isAdmin && (
          <div className="bg-white rounded-lg shadow-sm border border-gray-100 p-5 space-y-5">
            <h2 className="text-base font-semibold text-gray-800">Team Analytics</h2>

            {/* Insight stat cards */}
            <div className="grid grid-cols-4 gap-3">
              {[
                { label: "Avg Days on Bench",  value: stats.avgDays,       sub: "across active associates" },
                { label: "Total Submissions",   value: stats.totalSubmissions, sub: "client proposals made" },
                { label: "Selection Rate",      value: `${stats.selectionRate}%`, sub: `${stats.selected} of ${stats.totalSubmissions} placed` },
                { label: "Courses Tracked",     value: coursePool.length,   sub: "in the course pool" },
              ].map(s => (
                <div key={s.label} className="bg-indigo-50 rounded-lg p-3 border border-indigo-100">
                  <div className="text-2xl font-bold text-indigo-700">{s.value}</div>
                  <div className="text-xs font-semibold text-gray-600 mt-1">{s.label}</div>
                  <div className="text-xs text-gray-400 mt-0.5">{s.sub}</div>
                </div>
              ))}
            </div>

            {/* Charts */}
            <div className="grid grid-cols-2 gap-5">
              <div className="bg-gray-50 rounded-lg p-3"><canvas ref={chart1Ref} /></div>
              <div className="bg-gray-50 rounded-lg p-3"><canvas ref={chart2Ref} /></div>
            </div>

            {/* Proposals by project code */}
            {Object.keys(stats.byProject).length > 0 && (
              <div>
                <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Proposals by Project Code</h3>
                <div className="flex flex-wrap gap-2">
                  {Object.entries(stats.byProject)
                    .sort((a, b) => b[1] - a[1])
                    .map(([code, count]) => (
                      <span key={code} className="bg-indigo-50 text-indigo-700 border border-indigo-200 px-3 py-1.5 rounded-lg text-sm">
                        <span className="font-semibold">{code}</span>
                        <span className="ml-1.5 text-indigo-400">{count} proposed</span>
                      </span>
                    ))}
                </div>
              </div>
            )}

            {active.length === 0 && (
              <p className="text-sm text-gray-400 italic text-center py-4">No active associates — add data in Edit mode.</p>
            )}
          </div>
        )}

        {/* Admin Panel */}
        {isAdmin && (
          <div className="bg-white p-5 rounded-lg shadow-sm border border-gray-100 space-y-5">
            <div>
              <h2 className="text-sm font-semibold text-gray-700 mb-3">Manage Courses</h2>
              <div className="flex gap-2">
                <input
                  value={newCourse}
                  onChange={e => setNewCourse(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && addCourse()}
                  placeholder="New course name…"
                  className="border border-gray-300 rounded px-3 py-1.5 text-sm flex-1 max-w-xs focus:outline-none focus:ring-2 focus:ring-indigo-400"
                />
                <button onClick={addCourse} className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-1.5 rounded text-sm font-medium transition-colors">
                  Add Course
                </button>
              </div>
              <div className="flex flex-wrap gap-2 mt-3">
                {coursePool.map(c => (
                  <span key={c} className="bg-indigo-50 text-indigo-700 px-3 py-1 rounded-full flex items-center gap-1.5 text-sm border border-indigo-200">
                    {c}
                    <button onClick={() => removeCourse(c)} className="text-indigo-400 hover:text-red-500 transition-colors font-bold leading-none">×</button>
                  </span>
                ))}
              </div>
            </div>

            {archived.length > 0 && (
              <div className="pt-4 border-t border-gray-100">
                <h2 className="text-sm font-semibold text-gray-700 mb-3">Archived Associates <span className="text-gray-400 font-normal">({archived.length})</span></h2>
                <div className="space-y-2">
                  {archived.map(a => (
                    <div key={a.id} className="flex justify-between items-center bg-gray-50 px-4 py-2.5 rounded-lg border border-gray-100">
                      <div className="flex items-center gap-3">
                        <span className="font-medium text-sm text-gray-700">{a.name}</span>
                        <span className="text-xs text-gray-400">{days(a.joinDate)} days on bench</span>
                        {a.clients.some(c => c.outcome === "Selected") && (
                          <span className="text-xs bg-green-100 text-green-700 px-2 py-0.5 rounded-full font-medium">Placed</span>
                        )}
                        <span className="text-xs text-gray-400">{a.clients.length} client{a.clients.length !== 1 ? "s" : ""}</span>
                      </div>
                      <button
                        onClick={() => unarchiveAssociate(a.id)}
                        className="text-xs border border-gray-300 hover:border-indigo-400 hover:text-indigo-600 text-gray-500 px-3 py-1 rounded transition-all"
                      >
                        Restore
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {archived.length === 0 && (
              <p className="text-xs text-gray-400 italic pt-2 border-t border-gray-100">No archived associates yet. Archive placed or inactive associates via Edit mode.</p>
            )}
          </div>
        )}

        {/* Add Associate Form */}
        {!isAdmin && !managerView && (
          <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-100">
            <h2 className="text-sm font-semibold text-gray-700 mb-3">Add Associate</h2>
            <div className="flex gap-2 flex-wrap items-center">
              <input
                value={name}
                onChange={e => setName(e.target.value)}
                onKeyDown={e => e.key === "Enter" && addAssociate()}
                placeholder="Full name"
                className="border border-gray-300 rounded px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-400"
              />
              <input
                type="date"
                value={joinDate}
                onChange={e => setJoinDate(e.target.value)}
                className="border border-gray-300 rounded px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-400"
              />
              <button onClick={addAssociate} className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-1.5 rounded text-sm font-medium transition-colors">
                Add
              </button>
            </div>
          </div>
        )}

        {/* Associate List */}
        {!isAdmin && active.map(a => {
          const avg = avgCompletion(a);
          const placed = a.clients.some(c => c.outcome === "Selected");
          const barColor = avg >= 80 ? "bg-green-400" : avg >= 40 ? "bg-yellow-400" : avg > 0 ? "bg-orange-400" : "bg-gray-200";

          return (
            <div key={a.id} className="bg-white rounded-lg shadow-sm border border-gray-100 overflow-hidden">
              <div
                className="grid grid-cols-[240px_180px_1fr] px-4 py-3 items-center gap-3 cursor-pointer hover:bg-gray-50 transition-colors"
                onClick={() => !managerView && setExpandedId(expandedId === a.id ? null : a.id)}
              >
                {/* Name + indicator */}
                <div className="flex items-center gap-2.5">
                  <div className={`w-1.5 h-9 rounded-full flex-shrink-0 ${barColor}`} title={`Avg ${avg}%`} />
                  <div>
                    <div className="font-semibold text-gray-800">{a.name}</div>
                    <div className="text-xs text-gray-400">{days(a.joinDate)} days on bench</div>
                  </div>
                </div>

                {/* Pipeline status */}
                <div className="flex items-center gap-2">
                  {placed
                    ? <span className="text-xs bg-green-100 text-green-700 font-semibold px-2.5 py-1 rounded-full">Placed</span>
                    : a.clients.length > 0
                    ? <span className="text-xs bg-blue-100 text-blue-700 font-semibold px-2.5 py-1 rounded-full">{a.clients.length} proposed</span>
                    : <span className="text-xs text-gray-400 italic">No proposals yet</span>
                  }
                  {a.clients.some(c => c.internal === "Done" || c.client === "Done") && !placed && (
                    <span className="text-xs bg-yellow-100 text-yellow-700 px-2 py-0.5 rounded-full">In rounds</span>
                  )}
                </div>

                {/* Course chips */}
                <div className="flex gap-1.5 flex-wrap justify-end">
                  {Object.entries(a.courses).map(([k, v]) => (
                    <span key={k} className={`text-xs px-2 py-0.5 rounded-full font-medium ${courseChip(v)}`}>
                      {k}: {v}%
                    </span>
                  ))}
                </div>
              </div>

              {!managerView && expandedId === a.id && (
                <div className="border-t border-gray-100 bg-gray-50 p-4 space-y-5">

                  {/* Course sliders */}
                  <div>
                    <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2.5">Course Progress</h3>
                    <div className="space-y-2">
                      {coursePool.map(c => (
                        <div key={c} className="flex items-center gap-3">
                          <span className="text-sm text-gray-600 w-32 truncate flex-shrink-0">{c}</span>
                          <input
                            type="range" min="0" max="100"
                            value={a.courses[c] || 0}
                            onChange={e => updateCourse(a.id, c, e.target.value)}
                            className="flex-1 accent-indigo-600"
                          />
                          <span className={`text-xs px-2 py-0.5 rounded-full font-semibold w-12 text-center flex-shrink-0 ${courseChip(a.courses[c] || 0)}`}>
                            {a.courses[c] || 0}%
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Client pipeline */}
                  <div>
                    <div className="flex justify-between items-center mb-2.5">
                      <h3 className="text-xs font-semibold text-gray-500 uppercase tracking-wider">Client Pipeline</h3>
                      <button onClick={() => addClient(a.id)} className="text-xs bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1 rounded transition-colors">
                        + Add Client
                      </button>
                    </div>

                    {a.clients.length > 0 ? (
                      <div className="space-y-1.5">
                        <div className="grid grid-cols-[72px_1fr_110px_110px_100px_28px] gap-2 text-xs font-semibold text-gray-400 px-1 pb-1">
                          <span>Proj Code</span><span>Client</span><span>Internal Round</span><span>Client Round</span><span>Outcome</span><span />
                        </div>
                        {a.clients.map((c, j) => (
                          <div key={j} className="grid grid-cols-[72px_1fr_110px_110px_100px_28px] gap-2 items-center">
                            <input
                              value={c.projectCode || ""}
                              onChange={e => updateClient(a.id, j, "projectCode", e.target.value)}
                              placeholder="e.g. P001"
                              className="border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:ring-1 focus:ring-indigo-400"
                            />
                            <input
                              value={c.name}
                              onChange={e => updateClient(a.id, j, "name", e.target.value)}
                              placeholder="Client name"
                              className="border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:ring-1 focus:ring-indigo-400"
                            />
                            <select
                              value={c.internal}
                              onChange={e => updateClient(a.id, j, "internal", e.target.value)}
                              className={`border rounded px-1.5 py-1 text-xs focus:outline-none ${c.internal === "Done" ? "border-blue-300 bg-blue-50 text-blue-700" : "border-gray-200"}`}
                            >
                              <option>Not Started</option><option>Done</option>
                            </select>
                            <select
                              value={c.client}
                              onChange={e => updateClient(a.id, j, "client", e.target.value)}
                              className={`border rounded px-1.5 py-1 text-xs focus:outline-none ${c.client === "Done" ? "border-orange-300 bg-orange-50 text-orange-700" : "border-gray-200"}`}
                            >
                              <option>Not Started</option><option>Done</option>
                            </select>
                            <select
                              value={c.outcome}
                              onChange={e => updateClient(a.id, j, "outcome", e.target.value)}
                              className={`border rounded px-1.5 py-1 text-xs focus:outline-none ${
                                c.outcome === "Selected" ? "border-green-300 bg-green-50 text-green-700" :
                                c.outcome === "Rejected" ? "border-red-300 bg-red-50 text-red-700" :
                                "border-gray-200 text-gray-600"
                              }`}
                            >
                              <option>Pending</option><option>Selected</option><option>Rejected</option>
                            </select>
                            <button onClick={() => removeClient(a.id, j)} className="text-gray-300 hover:text-red-500 transition-colors text-lg leading-none text-center">×</button>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <p className="text-xs text-gray-400 italic">No clients added yet.</p>
                    )}
                  </div>

                  {/* Archive */}
                  <div className="pt-3 border-t border-gray-200 flex justify-end">
                    <button
                      onClick={() => archiveAssociate(a.id)}
                      className="text-xs border border-gray-300 hover:border-red-400 hover:text-red-600 text-gray-500 px-3 py-1.5 rounded transition-all"
                    >
                      Archive Associate
                    </button>
                  </div>
                </div>
              )}
            </div>
          );
        })}

        {/* Empty state */}
        {!isAdmin && active.length === 0 && (
          <div className="text-center py-16 text-gray-400">
            <p className="text-lg font-medium">{archived.length > 0 ? "All associates are archived." : "No associates on bench."}</p>
            <p className="text-sm mt-1">
              {managerView ? 'Switch to "Edit" mode to add associates.' : "Use the form above to get started."}
            </p>
          </div>
        )}

      </div>
    </div>
  );
}