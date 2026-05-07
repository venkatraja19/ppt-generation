import { useState, useRef } from "react";
import axios from "axios";
import './App.css';

const Spinner = () => (
  <svg className="spinner-icon" fill="none" viewBox="0 0 24 24">
    <circle className="spinner-track" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
    <path className="spinner-fill" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
  </svg>
);

export default function AIPPTGenerator() {
  const [topic, setTopic]       = useState("");
  const [numSlides, setNumSlides] = useState(10);
  const [loading, setLoading]   = useState(false);
  const [hasGenerated, setHasGenerated] = useState(false);
  const [error, setError]       = useState("");
  const inputRef = useRef(null);

  const handleGenerate = async () => {
    if (!topic.trim()) { inputRef.current?.focus(); return; }
    setLoading(true);
    setError("");
    try {
      const res = await axios.post(
        `${import.meta.env.VITE_API_URL || "http://127.0.0.1:5000"}/generate`,
        { topic, num_slides: numSlides },
        { responseType: "blob" }
      );
      const url  = window.URL.createObjectURL(new Blob([res.data]));
      const link = document.createElement("a");
      link.href  = url;
      link.download = `${topic.slice(0, 40).replace(/\s+/g, "_")}.pptx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
      setHasGenerated(true);
    } catch (err) {
      console.error(err);
      const msg = err?.response?.data?.error || err?.message || "Unknown error";
      setError(`Failed to generate PPT: ${msg}`);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app-wrapper">
      {/* Nav */}
      <div className="nav-bar">
        <div className="nav-inner">
          <div className="nav-brand">
            <div className="nav-logo">
              <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                <rect x="1" y="1" width="6" height="6" rx="1.5" fill="white" />
                <rect x="9" y="1" width="6" height="6" rx="1.5" fill="white" opacity="0.6" />
                <rect x="1" y="9" width="6" height="6" rx="1.5" fill="white" opacity="0.6" />
                <rect x="9" y="9" width="6" height="6" rx="1.5" fill="white" opacity="0.3" />
              </svg>
            </div>
            <span className="nav-title">SlideAI</span>
          </div>
          <span className="nav-powered">Powered by Groq</span>
        </div>
      </div>

      <div className="main-content">
        <div className="hero-text">
          <h1 className="hero-heading">
            Generate slides in seconds
          </h1>
          <p className="hero-subtext">
            Describe your presentation in detail — AI will craft structured, image-rich slides for you.
          </p>
        </div>

        {/* Input card */}
        <div className="input-card">
          <textarea
            ref={inputRef}
            value={topic}
            onChange={(e) => setTopic(e.target.value)}
            placeholder="e.g. Create a 12-slide presentation on AI in Healthcare covering diagnosis, drug discovery, patient monitoring, ethics, and future trends. Use a professional tone."
            rows={4}
            className="topic-textarea"
            disabled={loading}
          />
          <div className="card-footer">
            <div className="slides-control">
              <label className="slides-label">Slides</label>
              <input
                type="number"
                min={3}
                max={80}
                value={numSlides}
                onChange={(e) => {
                  const val = e.target.value;
                  if (val === "") {
                    setNumSlides("");
                    return;
                  }
                  const num = Number(val);
                  if (num >= 3 && num <= 80) {
                    setNumSlides(num);
                  } else {
                    setNumSlides(num);
                  }
                }}
                onBlur={() => {
                  if (numSlides === "" || numSlides < 3) setNumSlides(3);
                  if (numSlides > 80) setNumSlides(80);
                }}
                className="slides-input"
                disabled={loading}
              />
              <span className="slides-range">(3–80)</span>
            </div>
            <button
              onClick={handleGenerate}
              disabled={loading || !topic.trim()}
              className="generate-btn"
            >
              {loading ? (
                <>
                  <Spinner />
                  <span>Generating…</span>
                </>
              ) : (
                <>
                  <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                    <path d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5L8 1z" fill="currentColor" />
                  </svg>
                  <span>Generate</span>
                </>
              )}
            </button>
          </div>
        </div>

        {error && <p className="error-msg">{error}</p>}

        {loading && (
          <div className="loading-section">
            <div className="loading-dots-row">
              <div className="loading-dot dot-delay-0" />
              <div className="loading-dot dot-delay-1" />
              <div className="loading-dot dot-delay-2" />
              <span className="loading-text">Crafting your presentation…</span>
            </div>
            <div className="skeleton-grid">
              {Array.from({ length: 6 }).map((_, i) => (
                <div key={i} className="skeleton-card">
                  <div className="skeleton-line skeleton-line-short" />
                  <div className="skeleton-line skeleton-line-medium" />
                  <div className="skeleton-lines-group">
                    <div className="skeleton-line skeleton-line-full" />
                    <div className="skeleton-line skeleton-line-5of6" />
                    <div className="skeleton-line skeleton-line-4of6" />
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {!loading && hasGenerated && (
          <div className="result-section">
            <div className="result-icon-wrapper result-icon-green">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" className="icon-green">
                <path d="M5 13l4 4L19 7" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </div>
            <p className="result-title">Your presentation was downloaded!</p>
            <p className="result-subtitle">Check your downloads folder for the .pptx file.</p>
            <button
              onClick={() => { setHasGenerated(false); setTopic(""); }}
              className="generate-another-btn"
            >
              Generate another
            </button>
          </div>
        )}

        {!loading && !hasGenerated && (
          <div className="result-section">
            <div className="result-icon-wrapper result-icon-indigo">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" className="icon-indigo">
                <rect x="3" y="3" width="8" height="8" rx="2" fill="currentColor" />
                <rect x="13" y="3" width="8" height="8" rx="2" fill="currentColor" opacity="0.5" />
                <rect x="3" y="13" width="8" height="8" rx="2" fill="currentColor" opacity="0.5" />
                <rect x="13" y="13" width="8" height="8" rx="2" fill="currentColor" opacity="0.25" />
              </svg>
            </div>
            <p className="result-placeholder-text">Your slides will appear here</p>
          </div>
        )}
      </div>
    </div>
  );
}


// import { useState, useRef } from "react";
// import axios from "axios";

// const Spinner = () => (
//   <svg className="animate-spin h-4 w-4" fill="none" viewBox="0 0 24 24">
//     <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
//     <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
//   </svg>
// );

// export default function AIPPTGenerator() {
//   const [topic, setTopic]       = useState("");
//   const [numSlides, setNumSlides] = useState(10);
//   const [loading, setLoading]   = useState(false);
//   const [hasGenerated, setHasGenerated] = useState(false);
//   const [error, setError]       = useState("");
//   const inputRef = useRef(null);

//   const handleGenerate = async () => {
//     if (!topic.trim()) { inputRef.current?.focus(); return; }
//     setLoading(true);
//     setError("");
//     try {
//       const res = await axios.post(
//         `${import.meta.env.VITE_API_URL || "http://127.0.0.1:5000"}/generate`,
//         { topic, num_slides: numSlides },
//         { responseType: "blob" }
//       );
//       const url  = window.URL.createObjectURL(new Blob([res.data]));
//       const link = document.createElement("a");
//       link.href  = url;
//       link.download = `${topic.slice(0, 40).replace(/\s+/g, "_")}.pptx`;
//       document.body.appendChild(link);
//       link.click();
//       link.remove();
//       window.URL.revokeObjectURL(url);
//       setHasGenerated(true);
//     } catch (err) {
//       console.error(err);
//       setError("Failed to generate PPT. Is the Flask backend running on port 5000?");
//     } finally {
//       setLoading(false);
//     }
//   };

//   return (
//     <div className="min-h-screen bg-slate-50 font-sans">
//       {/* Nav */}
//       <div className="border-b border-slate-200 bg-white">
//         <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
//           <div className="flex items-center gap-2.5">
//             <div className="w-7 h-7 rounded-lg bg-indigo-600 flex items-center justify-center">
//               <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
//                 <rect x="1" y="1" width="6" height="6" rx="1.5" fill="white" />
//                 <rect x="9" y="1" width="6" height="6" rx="1.5" fill="white" opacity="0.6" />
//                 <rect x="1" y="9" width="6" height="6" rx="1.5" fill="white" opacity="0.6" />
//                 <rect x="9" y="9" width="6" height="6" rx="1.5" fill="white" opacity="0.3" />
//               </svg>
//             </div>
//             <span className="text-sm font-semibold text-slate-800 tracking-tight">SlideAI</span>
//           </div>
//           <span className="text-xs text-slate-400 font-medium">Powered by Groq</span>
//         </div>
//       </div>

//       <div className="max-w-3xl mx-auto px-6 py-12">
//         <div className="text-center mb-10">
//           <h1 className="text-4xl font-bold text-slate-900 tracking-tight mb-3">
//             Generate slides in seconds
//           </h1>
//           <p className="text-slate-500 text-base max-w-md mx-auto">
//             Describe your presentation in detail — AI will craft structured, image-rich slides for you.
//           </p>
//         </div>

//         {/* Input card */}
//         <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 mb-3">
//           <textarea
//             ref={inputRef}
//             value={topic}
//             onChange={(e) => setTopic(e.target.value)}
//             placeholder="e.g. Create a 12-slide presentation on AI in Healthcare covering diagnosis, drug discovery, patient monitoring, ethics, and future trends. Use a professional tone."
//             rows={4}
//             className="w-full text-sm text-slate-800 placeholder-slate-400 bg-transparent outline-none resize-none"
//             disabled={loading}
//           />
//           <div className="flex items-center justify-between mt-4 pt-4 border-t border-slate-100">
//             <div className="flex items-center gap-3">
//               <label className="text-xs text-slate-500 font-medium">Slides</label>
//               {/* <input
//                 type="number"
//                 min={3} max={80}
//                 value={numSlides}
//                 onChange={(e) => setNumSlides(Math.min(80, Math.max(3, Number(e.target.value))))}
//                 className="w-16 text-sm text-center border border-slate-200 rounded-lg px-2 py-1 outline-none focus:border-indigo-400"
//                 disabled={loading}              /> */}
//                 <input
//   type="number"
//   min={3}
//   max={80}
//   value={numSlides}
//   onChange={(e) => {
//     const val = e.target.value;

//     // allow empty while typing
//     if (val === "") {
//       setNumSlides("");
//       return;
//     }

//     const num = Number(val);

//     // allow typing within range only
//     if (num >= 3 && num <= 80) {
//       setNumSlides(num);
//     } else {
//       setNumSlides(num); // still allow typing (like 1, then 10)
//     }
//   }}
//   onBlur={() => {
//     // fix value only when user leaves input
//     if (numSlides === "" || numSlides < 3) setNumSlides(3);
//     if (numSlides > 80) setNumSlides(80);
//   }}
//   className="w-16 text-sm text-center border border-slate-200 rounded-lg px-2 py-1 outline-none focus:border-indigo-400"
//   disabled={loading}
// />
//               <span className="text-xs text-slate-400">(3–80)</span>
//             </div>
//             <button
//               onClick={handleGenerate}
//               disabled={loading || !topic.trim()}
//               className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 text-white text-sm font-semibold rounded-xl hover:bg-indigo-700 active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-150"
//             >
//               {loading ? (<><Spinner /><span>Generating…</span></>) : (
//                 <><svg width="14" height="14" viewBox="0 0 16 16" fill="none">
//                   <path d="M8 1l1.5 4.5L14 7l-4.5 1.5L8 13l-1.5-4.5L2 7l4.5-1.5L8 1z" fill="currentColor" />
//                 </svg><span>Generate</span></>
//               )}
//             </button>
//           </div>
//         </div>

//         {error && <p className="text-center text-sm text-red-500 mb-6">{error}</p>}

//         {loading && (
//           <div className="mt-10">
//             <div className="flex items-center justify-center gap-2 mb-6">
//               <div className="w-1.5 h-1.5 rounded-full bg-indigo-400 animate-bounce [animation-delay:0ms]" />
//               <div className="w-1.5 h-1.5 rounded-full bg-indigo-400 animate-bounce [animation-delay:150ms]" />
//               <div className="w-1.5 h-1.5 rounded-full bg-indigo-400 animate-bounce [animation-delay:300ms]" />
//               <span className="text-sm text-slate-500 ml-2">Crafting your presentation…</span>
//             </div>
//             <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
//               {Array.from({ length: 6 }).map((_, i) => (
//                 <div key={i} className="bg-white rounded-2xl border border-slate-200 p-5 animate-pulse">
//                   <div className="h-3 bg-slate-100 rounded w-1/4 mb-4" />
//                   <div className="h-4 bg-slate-200 rounded w-3/4 mb-4" />
//                   <div className="space-y-2">
//                     <div className="h-3 bg-slate-100 rounded w-full" />
//                     <div className="h-3 bg-slate-100 rounded w-5/6" />
//                     <div className="h-3 bg-slate-100 rounded w-4/6" />
//                   </div>
//                 </div>
//               ))}
//             </div>
//           </div>
//         )}

//         {!loading && hasGenerated && (
//           <div className="mt-16 text-center">
//             <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl bg-green-50 mb-4">
//               <svg width="24" height="24" viewBox="0 0 24 24" fill="none" className="text-green-500">
//                 <path d="M5 13l4 4L19 7" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
//               </svg>
//             </div>
//             <p className="text-sm font-semibold text-slate-700">Your presentation was downloaded!</p>
//             <p className="text-xs text-slate-400 mt-1">Check your downloads folder for the .pptx file.</p>
//             <button
//               onClick={() => { setHasGenerated(false); setTopic(""); }}
//               className="mt-4 text-xs font-medium text-indigo-600 hover:underline"
//             >
//               Generate another
//             </button>
//           </div>
//         )}

//         {!loading && !hasGenerated && (
//           <div className="mt-16 text-center">
//             <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl bg-indigo-50 mb-4">
//               <svg width="24" height="24" viewBox="0 0 24 24" fill="none" className="text-indigo-400">
//                 <rect x="3" y="3" width="8" height="8" rx="2" fill="currentColor" />
//                 <rect x="13" y="3" width="8" height="8" rx="2" fill="currentColor" opacity="0.5" />
//                 <rect x="3" y="13" width="8" height="8" rx="2" fill="currentColor" opacity="0.5" />
//                 <rect x="13" y="13" width="8" height="8" rx="2" fill="currentColor" opacity="0.25" />
//               </svg>
//             </div>
//             <p className="text-sm text-slate-400">Your slides will appear here</p>
//           </div>
//         )}
//       </div>
//     </div>
//   );
// }


