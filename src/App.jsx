import { useRef, useState } from "react";
import * as XLSX from "xlsx";

const SAMPLE = [
  { id: "contact-1", name: "Aarav Sharma", email: "aarav@example.com", phone: "9876543210" },
  { id: "contact-2", name: "Priya Das", email: "priya@example.com", phone: "9123456780" }
];
const HEADERS = ["id", "name", "email", "phone"];

export default function App() {
  const [contacts, setContacts] = useState(SAMPLE);
  const [form, setForm] = useState({ name: "", email: "", phone: "" });
  const [editingId, setEditingId] = useState(null);
  const [search, setSearch] = useState("");
  const [message, setMessage] = useState("Connect an Excel workbook to load and save your contacts.");
  const [fileName, setFileName] = useState("");
  const [fileHandle, setFileHandle] = useState(null);
  const [canWriteDirectly, setCanWriteDirectly] = useState(
    typeof window !== "undefined" && "showOpenFilePicker" in window
  );
  const importInput = useRef(null);

  function parseWorkbook(buffer) {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets["Contacts"] || workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, { defval: "" }).map((row, index) => ({
      id: String(row.id || `contact-${Date.now()}-${index}`),
      name: String(row.name || ""),
      email: String(row.email || ""),
      phone: String(row.phone || "")
    }));
  }

  async function connectSpreadsheet() {
    try {
      if ("showOpenFilePicker" in window) {
        const [handle] = await window.showOpenFilePicker({
          multiple: false,
          types: [{ description: "Excel spreadsheet", accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            "application/vnd.ms-excel": [".xls"]
          }}]
        });
        const file = await handle.getFile();
        const rows = parseWorkbook(await file.arrayBuffer());
        setFileHandle(handle);
        setFileName(file.name);
        setContacts(rows);
        setMessage(`Connected to ${file.name}. Make changes and click Save to spreadsheet.`);
        setCanWriteDirectly(true);
      } else {
        importInput.current?.click();
      }
    } catch (error) {
      if (error?.name !== "AbortError") setMessage("Could not open the spreadsheet. Please try again.");
    }
  }

  async function importSpreadsheet(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setContacts(parseWorkbook(await file.arrayBuffer()));
      setFileName(file.name);
      setFileHandle(null);
      setMessage(`Imported ${file.name}. After editing, download the updated spreadsheet.`);
    } catch {
      setMessage("Could not read this file. Please select a valid Excel workbook.");
    }
    event.target.value = "";
  }

  function makeWorkbook(rows) {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows.map(({ id, name, email, phone }) => ({ id, name, email, phone })), { header: HEADERS });
    XLSX.utils.book_append_sheet(workbook, sheet, "Contacts");
    return workbook;
  }

  async function saveSpreadsheet() {
    try {
      const workbook = makeWorkbook(contacts);
      if (fileHandle && "createWritable" in fileHandle) {
        const writable = await fileHandle.createWritable();
        await writable.write(XLSX.write(workbook, { bookType: "xlsx", type: "array" }));
        await writable.close();
        setMessage(`Saved ${contacts.length} contacts to ${fileName}.`);
      } else {
        XLSX.writeFile(workbook, "contacts-updated.xlsx");
        setMessage("Downloaded contacts-updated.xlsx. Replace your original file manually if needed.");
      }
    } catch (error) {
      console.error(error);
      setMessage("Could not save the spreadsheet. Try reconnecting the file.");
    }
  }

  function updateField(event) {
    setForm({ ...form, [event.target.name]: event.target.value });
  }

  function submitForm(event) {
    event.preventDefault();
    const name = form.name.trim();
    const email = form.email.trim();
    if (!name || !email) return;
    if (editingId) {
      setContacts((current) => current.map((item) => item.id === editingId
        ? { ...item, name, email, phone: form.phone.trim() } : item));
      setMessage("Contact updated in the app. Click Save to spreadsheet to persist the change.");
    } else {
      setContacts((current) => [...current, {
        id: globalThis.crypto?.randomUUID?.() || `contact-${Date.now()}`,
        name, email, phone: form.phone.trim()
      }]);
      setMessage("Contact added in the app. Click Save to spreadsheet to persist the change.");
    }
    setForm({ name: "", email: "", phone: "" });
    setEditingId(null);
  }

  function editContact(contact) {
    setForm({ name: contact.name, email: contact.email, phone: contact.phone });
    setEditingId(contact.id);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function deleteContact(id) {
    if (!window.confirm("Delete this contact?")) return;
    setContacts((current) => current.filter((item) => item.id !== id));
    setMessage("Contact deleted in the app. Click Save to spreadsheet to persist the change.");
  }

  const filtered = contacts.filter((contact) =>
    [contact.name, contact.email, contact.phone].some((value) =>
      String(value ?? "").toLowerCase().includes(search.toLowerCase())
    )
  );

  return (
    <main className="page">
      <header className="hero">
        <div className="eyebrow">REACT + JAVASCRIPT + EXCEL</div>
        <h1>Contact Manager</h1>
        <p>A simple contact app that reads and saves an Excel workbook directly in your browser.</p>
      </header>

      <section className="panel connection">
        <div className="connection-copy">
          <h2>Spreadsheet</h2>
          <p className="muted">{fileName ? `Selected file: ${fileName}` : "No spreadsheet connected yet."}</p>
        </div>
        <div className="connection-actions">
          <button className="secondary" onClick={connectSpreadsheet}>Connect spreadsheet</button>
          <button className="primary" onClick={saveSpreadsheet}>Save to spreadsheet</button>
          <input ref={importInput} type="file" accept=".xlsx,.xls" hidden onChange={importSpreadsheet} />
        </div>
        <p className="status" role="status">{message}</p>
      </section>

      <section className="panel">
        <h2>{editingId ? "Edit contact" : "Add a contact"}</h2>
        <form onSubmit={submitForm} className="contact-form">
          <label>Full name *
            <input name="name" value={form.name} onChange={updateField} placeholder="e.g. Riya Patnaik" required />
          </label>
          <label>Email address *
            <input name="email" type="email" value={form.email} onChange={updateField} placeholder="riya@example.com" required />
          </label>
          <label>Phone number
            <input name="phone" value={form.phone} onChange={updateField} placeholder="Optional" />
          </label>
          <div className="form-actions">
            <button className="primary" type="submit">{editingId ? "Save changes" : "Add contact"}</button>
            {editingId && <button className="secondary" type="button" onClick={() => {
              setEditingId(null); setForm({ name: "", email: "", phone: "" });
            }}>Cancel</button>}
          </div>
        </form>
      </section>

      <section className="panel">
        <div className="list-heading">
          <div><h2>Contacts</h2><p className="muted">{contacts.length} total contacts</p></div>
          <input className="search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search contacts..." aria-label="Search contacts" />
        </div>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Name</th><th>Email</th><th>Phone</th><th>Actions</th></tr></thead>
            <tbody>
              {filtered.map((contact) => <tr key={contact.id}>
                <td><strong>{contact.name}</strong></td><td>{contact.email}</td><td>{contact.phone || "—"}</td>
                <td className="actions">
                  <button className="text-button" onClick={() => editContact(contact)}>Edit</button>
                  <button className="text-button danger" onClick={() => deleteContact(contact.id)}>Delete</button>
                </td>
              </tr>)}
              {filtered.length === 0 && <tr><td colSpan="4" className="empty">No contacts found.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
      <footer>Browser-only spreadsheet demo · {canWriteDirectly ? "Direct file saving supported by this browser" : "Import and download mode"}</footer>
    </main>
  );
}
