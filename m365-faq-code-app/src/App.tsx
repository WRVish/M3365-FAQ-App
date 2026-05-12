import { useEffect, useMemo, useState } from "react";
import "./App.css";
import { M365_FAQsService } from "./generated/services/M365_FAQsService";
import { fieldMap } from "./fieldMap";
import { appConfig } from "./config";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { Copy, Check, Menu, X, SearchX, Paperclip } from "lucide-react";

type Attachment = { id: string; url: string; name: string };

type FaqItem = {
  id: string;
  title: string;
  answer: string;
  category: string;
  subCategory: string;
  active: boolean;
  pinned: boolean;
  attachments: Attachment[];
};

function getText(value: unknown, fallback = "") {
  if (value === null || value === undefined) return fallback;
  if (typeof value === "string") return value;
  if (typeof value === "object" && value !== null && "Value" in value) {
    return String((value as { Value?: unknown }).Value ?? fallback);
  }
  return String(value);
}

function findField(raw: Record<string, unknown>, candidates: string[]) {
  const normalized: Record<string, unknown> = {};
  Object.keys(raw).forEach((key) => {
    normalized[key.toLowerCase()] = raw[key];
  });

  for (const candidate of candidates) {
    if (candidate in raw) return raw[candidate];
    const lower = candidate.toLowerCase();
    if (lower in normalized) return normalized[lower];
  }

  return undefined;
}

function getBoolean(value: unknown) {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    return value.toLowerCase() === "yes" || value.toLowerCase() === "true";
  }
  return Boolean(value);
}

function normalizeFaq(item: unknown): FaqItem {
  const raw = item as Record<string, unknown>;

  const attachments: Attachment[] = [];
  const rawAttachments = raw["{Attachments}"] || raw["Attachments"] || findField(raw, ["{Attachments}", "Attachments", "attachments"]);
  
  if (Array.isArray(rawAttachments)) {
    for (const att of rawAttachments) {
      if (att && typeof att === "object") {
        const attObj = att as Record<string, any>;
        attachments.push({
          id: String(attObj.Id || attObj.id || ""),
          url: String(attObj.AbsoluteUri || attObj.absoluteUri || attObj.Url || attObj.url || attObj.Link || attObj.ServerRelativeUrl || ""),
          name: String(attObj.DisplayName || attObj.displayName || attObj.Name || attObj.name || attObj.FileName || "Attachment"),
        });
      }
    }
  }

  const hasAttachments = getBoolean(raw["{HasAttachments}"] || raw["HasAttachments"] || raw["hasAttachments"]);
  
  if (attachments.length === 0 && hasAttachments) {
    const id = String(raw["ID"] || raw["Id"] || raw["id"] || "");
    
    // Hardcoded workaround for item 1 as requested since SP hides the filename
    if (id === "1" || id === "1_.000") {
      attachments.push({
        id: "1-hardcoded",
        url: "https://vishtechtalk.sharepoint.com/sites/DemoSite/Lists/M365_FAQs/Attachments/1/M365_FAQ_List_With_Schema.csv?web=1",
        name: "M365_FAQ_List_With_Schema.csv"
      });
    } else {
      const link = String(raw["{Link}"] || `https://vishtechtalk.sharepoint.com/sites/DemoSite/Lists/M365%20FAQs/DispForm.aspx?ID=${id}`);
      attachments.push({
        id: "fallback",
        url: link,
        name: "View Attachments on SharePoint"
      });
    }
  }

  return {
    id: getText(findField(raw, fieldMap.id)),
    title: getText(findField(raw, fieldMap.title), "Untitled FAQ"),
    answer: getText(findField(raw, fieldMap.answer)),
    category: getText(findField(raw, fieldMap.category), "Uncategorized"),
    subCategory: getText(findField(raw, fieldMap.subCategory), "General"),
    active: getBoolean(findField(raw, fieldMap.active)),
    pinned: getBoolean(findField(raw, fieldMap.pinned)),
    attachments,
  };
}

function Highlight({ text, query }: { text: string; query: string }) {
  if (!query) return <>{text}</>;
  
  const regex = new RegExp(`(${query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, "gi");
  const parts = text.split(regex);
  
  return (
    <>
      {parts.map((part, i) =>
        regex.test(part) ? <mark key={i}>{part}</mark> : <span key={i}>{part}</span>
      )}
    </>
  );
}

function App() {
  const [faqs, setFaqs] = useState<FaqItem[]>([]);
  const [selectedCategory, setSelectedCategory] = useState("");
  const [selectedSubCategory, setSelectedSubCategory] = useState("");
  const [selectedFaq, setSelectedFaq] = useState<FaqItem | null>(null);
  const [expandedFaqId, setExpandedFaqId] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [uiMode, setUiMode] = useState<"accordion" | "modal">(appConfig.uiMode);

  useEffect(() => {
    async function loadFaqs() {
      try {
        setLoading(true);
        setError("");

        const result = await M365_FAQsService.getAll();
        const records = Array.isArray(result.data) ? result.data : [];
        const activeFaqs = records.map(normalizeFaq).filter((faq) => faq.active);

        setFaqs(activeFaqs);
        setSelectedCategory(activeFaqs[0]?.category ?? "");
      } catch (err) {
        console.error(err);
        setError("Unable to load FAQ records from SharePoint.");
      } finally {
        setLoading(false);
      }
    }

    loadFaqs();
  }, []);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setSelectedFaq(null);
      }
    };

    if (selectedFaq) {
      window.addEventListener("keydown", handleKeyDown);
    }

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [selectedFaq]);

  const categories = useMemo(() => {
    return Array.from(new Set(faqs.map((faq) => faq.category))).sort();
  }, [faqs]);

  const subCategories = useMemo(() => {
    return Array.from(
      new Set(
        faqs
          .filter((faq) => faq.category === selectedCategory)
          .map((faq) => faq.subCategory)
      )
    ).sort();
  }, [faqs, selectedCategory]);

  const filteredFaqs = useMemo(() => {
    const query = search.trim().toLowerCase();

    const results = faqs.filter((faq) => {
      const matchesCategory = faq.category === selectedCategory;
      const matchesSubCategory =
        !selectedSubCategory || faq.subCategory === selectedSubCategory;
      const matchesSearch =
        !query ||
        faq.title.toLowerCase().includes(query) ||
        faq.answer.toLowerCase().includes(query);

      return matchesCategory && matchesSubCategory && matchesSearch;
    });

    return results.sort((a, b) => {
      if (a.pinned !== b.pinned) {
        return a.pinned ? -1 : 1;
      }
      return a.title.localeCompare(b.title);
    });
  }, [faqs, selectedCategory, selectedSubCategory, search]);

  function handleCategorySelect(category: string) {
    setSelectedCategory(category);
    setSelectedSubCategory("");
    setSearch("");
    setCurrentPage(1);
    setExpandedFaqId(null);
    setIsSidebarOpen(false);
  }

  function handleFaqClick(faq: FaqItem) {
    if (uiMode === "accordion") {
      setExpandedFaqId((prev) => (prev === faq.id ? null : faq.id));
    } else {
      setSelectedFaq(faq);
    }
  }

  async function handleCopy(faq: FaqItem) {
    const attachmentsText = faq.attachments.length > 0 
      ? `\n\nAttachments:\n${faq.attachments.map(a => `- ${a.name}: ${a.url}`).join("\n")}`
      : "";
    
    const textToCopy = `**Q: ${faq.title}**\n\n${faq.answer}${attachmentsText}`;
    
    try {
      await navigator.clipboard.writeText(textToCopy);
      setCopiedId(faq.id);
      setTimeout(() => setCopiedId(null), 2000);
    } catch (err) {
      console.error("Failed to copy", err);
    }
  }

  const totalPages = Math.ceil(filteredFaqs.length / appConfig.maxItemsPerPage);
  const paginatedFaqs = filteredFaqs.slice(
    (currentPage - 1) * appConfig.maxItemsPerPage,
    currentPage * appConfig.maxItemsPerPage
  );

  const renderAnswerContent = (faq: FaqItem) => (
    <div className="answer-content">
      <div className="answer-actions">
        <button className="copy-btn" onClick={(e) => { e.stopPropagation(); handleCopy(faq); }}>
          {copiedId === faq.id ? <Check size={16} /> : <Copy size={16} />}
          {copiedId === faq.id ? "Copied!" : "Copy"}
        </button>
      </div>
      <div className="markdown-body">
        <ReactMarkdown remarkPlugins={[remarkGfm]}>{faq.answer}</ReactMarkdown>
      </div>
      {faq.attachments.length > 0 && (
        <div className="attachments-list">
          <strong><Paperclip size={16} /> Attachments</strong>
          <ul>
            {faq.attachments.map(att => (
              <li key={att.id}>
                <a href={att.url} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()}>
                  {att.name}
                </a>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );

  return (
    <main className="app-shell">
      <header className="app-header">
        <div className="brand">
          <button className="mobile-menu-btn" onClick={() => setIsSidebarOpen(!isSidebarOpen)}>
            {isSidebarOpen ? <X size={24} /> : <Menu size={24} />}
          </button>
          <h1>Microsoft 365 FAQ Center</h1>
        </div>

        <div className="header-actions">
          <input
            className="search"
            value={search}
            onChange={(event) => {
              setSearch(event.target.value);
              setCurrentPage(1);
            }}
            placeholder="Search FAQ"
            aria-label="Search FAQ"
          />
          <div className="view-toggle" role="group" aria-label="View mode">
            <button
              type="button"
              className={uiMode === "accordion" ? "active" : ""}
              onClick={() => { setUiMode("accordion"); setSelectedFaq(null); setExpandedFaqId(null); }}
              title="Accordion view"
            >
              ☰ <span className="toggle-label">Accordion</span>
            </button>
            <button
              type="button"
              className={uiMode === "modal" ? "active" : ""}
              onClick={() => { setUiMode("modal"); setExpandedFaqId(null); }}
              title="Modal view"
            >
              ⊞ <span className="toggle-label">Modal</span>
            </button>
          </div>
        </div>
      </header>

      {isSidebarOpen && <div className="sidebar-overlay" onClick={() => setIsSidebarOpen(false)} />}
      
      <aside className={`sidebar ${isSidebarOpen ? "open" : ""}`}>
        <p className="sidebar-heading">Categories</p>
        <nav className="category-list" aria-label="FAQ categories">
          {categories.map((category) => (
            <button
              key={category}
              type="button"
              className={category === selectedCategory ? "active" : ""}
              onClick={() => handleCategorySelect(category)}
            >
              {category}
            </button>
          ))}
        </nav>
      </aside>

      <section className="content">
        <header className="content-header">
          <div>
            <h2>{selectedCategory || "All FAQs"}</h2>
          </div>
        </header>

        <div className="tabs" role="tablist" aria-label="Sub categories">
          <button
            type="button"
            className={!selectedSubCategory ? "active" : ""}
            onClick={() => {
              setSelectedSubCategory("");
              setCurrentPage(1);
            }}
          >
            All
          </button>

          {subCategories.map((subCategory) => (
            <button
              key={subCategory}
              type="button"
              className={subCategory === selectedSubCategory ? "active" : ""}
              onClick={() => {
                setSelectedSubCategory(subCategory);
                setCurrentPage(1);
              }}
            >
              {subCategory}
            </button>
          ))}
        </div>

        {loading && <div className="state">Loading FAQs...</div>}
        {error && <div className="state error">{error}</div>}

        {!loading && !error && (
          <>
            <div className="faq-list">
              {paginatedFaqs.length === 0 ? (
                <div className="empty-state">
                  <SearchX size={48} className="empty-icon" />
                  <h3>No FAQs Found</h3>
                  <p>We couldn't find anything matching your search.</p>
                  <button className="clear-btn" onClick={() => setSearch("")}>
                    Clear Search
                  </button>
                </div>
              ) : (
                paginatedFaqs.map((faq) => {
                  const isExpanded = expandedFaqId === faq.id;
                  return (
                    <div key={faq.id || faq.title} className={`faq-item-container ${isExpanded ? "expanded" : ""}`}>
                      <button
                        type="button"
                        className="faq-item"
                        onClick={() => handleFaqClick(faq)}
                      >
                        <div className="faq-item-main">
                          <strong><Highlight text={faq.title} query={search} /></strong>
                          {faq.pinned && <span className="badge">★</span>}
                        </div>
                      </button>
                      {uiMode === "accordion" && isExpanded && (
                        <div className="faq-accordion-content">
                          {renderAnswerContent(faq)}
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>

            <div
              className="pagination"
              style={{ visibility: totalPages > 1 ? "visible" : "hidden" }}
            >
              <button
                type="button"
                disabled={currentPage === 1}
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              >
                Previous
              </button>
              <div className="page-numbers">
                {Array.from({ length: Math.max(1, totalPages) }).map((_, idx) => {
                  const page = idx + 1;
                  return (
                    <button
                      key={page}
                      type="button"
                      className={currentPage === page ? "active" : ""}
                      onClick={() => setCurrentPage(page)}
                    >
                      {page}
                    </button>
                  );
                })}
              </div>
              <button
                type="button"
                disabled={currentPage === totalPages || totalPages === 0}
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              >
                Next
              </button>
            </div>
          </>
        )}
      </section>

      {selectedFaq && uiMode === "modal" && (
        <div className="modal-backdrop" onClick={() => setSelectedFaq(null)}>
          <section
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="faq-title"
            onClick={(event) => event.stopPropagation()}
          >
            <header>
              <div>
                <p>
                  {selectedFaq.category} / {selectedFaq.subCategory}
                </p>
                <h3 id="faq-title"><Highlight text={selectedFaq.title} query={search} /></h3>
              </div>

              <button
                type="button"
                className="close"
                aria-label="Close FAQ details"
                onClick={() => setSelectedFaq(null)}
              >
                <X size={20} />
              </button>
            </header>

            {renderAnswerContent(selectedFaq)}
          </section>
        </div>
      )}
    </main>
  );
}

export default App;
