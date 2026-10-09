import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { loadContacts, saveContacts } from "./contactsDb.js";

const SAMPLE = [
  { id: "contact-1", name: "Aarav Sharma", email: "aarav@example.com", phone: "9876543210" },
  { id: "contact-2", name: "Priya Das", email: "priya@example.com", phone: "9123456780" }
];

export default function App() {
  const [contacts, setContacts] = useState([]);
  const [form, setForm] = useState({ name: "", email: "", phone: "" });
  const [editingId, setEditingId] = useState(null);
  const [search, setSearch] = useState("");
  const [message, setMessage] = useState("Opening the contacts database…");
  const [isLoaded, setIsLoaded] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [hasImportedChanges, setHasImportedChanges] = useState(false);
  const importInput = useRef(null);

  useEffect(() => {
    let isMounted = true;

    async function initializeDatabase() {
      try {
        let storedContacts = await loadContacts();
        if (storedContacts.length === 0) {
          storedContacts = SAMPLE;
          await saveContacts(storedContacts);
        }
        if (isMounted) {
          setContacts(storedContacts);
          setMessage("Contacts are saved in this browser's database.");
        }
      } catch (error) {
        console.error(error);
        if (isMounted) setMessage("Could not open the contacts database in this browser.");
      } finally {
        if (isMounted) setIsLoaded(true);
      }
    }

    initializeDatabase();
    return () => { isMounted = false; };
  }, []);

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

  async function importSpreadsheet(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setContacts(parseWorkbook(await file.arrayBuffer()));
      setHasImportedChanges(true);
      setMessage(`Loaded ${file.name}. Save once to store these contacts in the database.`);
    } catch {
      setMessage("Could not read this file. Please select a valid Excel workbook.");
    }
    event.target.value = "";
  }

  async function persistContacts(nextContacts, successMessage) {
    setIsSaving(true);
    try {
      await saveContacts(nextContacts);
      setContacts(nextContacts);
      setHasImportedChanges(false);
      setMessage(successMessage);
      return true;
    } catch (error) {
      console.error(error);
      setMessage("Could not save to the database. Please try again.");
      return false;
    } finally {
      setIsSaving(false);
    }
  }

  async function saveAllContacts() {
    await persistContacts(contacts, `Saved ${contacts.length} contacts to this browser's database.`);
  }

  function updateField(event) {
    setForm({ ...form, [event.target.name]: event.target.value });
  }

  async function submitForm(event) {
    event.preventDefault();
    const name = form.name.trim();
    const email = form.email.trim();
    if (!name || !email || isSaving) return;

    const nextContacts = editingId
      ? contacts.map((item) => item.id === editingId
        ? { ...item, name, email, phone: form.phone.trim() } : item)
      : [...contacts, {
        id: globalThis.crypto?.randomUUID?.() || `contact-${Date.now()}`,
        name,
        email,
        phone: form.phone.trim()
      }];
    const saved = await persistContacts(
      nextContacts,
      editingId ? "Contact updated in the database." : "Contact added to the database."
    );
    if (saved) {
      setForm({ name: "", email: "", phone: "" });
      setEditingId(null);
    }
  }

  function editContact(contact) {
    setForm({ name: contact.name, email: contact.email, phone: contact.phone });
    setEditingId(contact.id);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function deleteContact(id) {
    if (!window.confirm("Delete this contact?")) return;
    const nextContacts = contacts.filter((item) => item.id !== id);
    await persistContacts(nextContacts, "Contact deleted from the database.");
  }

  function cancelEdit() {
    setEditingId(null);
    setForm({ name: "", email: "", phone: "" });
  }

  const filtered = contacts.filter((contact) =>
    [contact.name, contact.email, contact.phone].some((value) =>
      String(value ?? "").toLowerCase().includes(search.toLowerCase())
    )
  );

  return (
    <main className="page">
      <header className="hero">
        <div className="eyebrow">REACT + JAVASCRIPT + DATABASE</div>
        <h1>Contact Manager</h1>
        <p>Add, edit, and keep your contacts in this browser without downloading a spreadsheet after every change.</p>
      </header>

      <section className="panel connection">
        <div className="connection-copy">
          <h2>Database</h2>
          <p className="muted">Contacts are saved in this browser and loaded automatically when you return.</p>
        </div>
        <div className="connection-actions">
          <button className="secondary" type="button" onClick={() => importInput.current?.click()} disabled={!isLoaded || isSaving}>
            Import Excel workbook
          </button>
          <button className="primary" type="button" onClick={saveAllContacts} disabled={!isLoaded || isSaving}>
            {isSaving ? "Saving…" : hasImportedChanges ? "Save imported contacts to database" : "Save to database"}
          </button>
          <input ref={importInput} type="file" accept=".xlsx,.xls" hidden onChange={importSpreadsheet} />
        </div>
        <p className="status" role="status">{message}</p>
      </section>

      <section className="panel">
        <h2>{editingId ? "Edit contact" : "Add a contact"}</h2>
        <form onSubmit={submitForm} className="contact-form">
          <label>Full name *
            <input name="name" value={form.name} onChange={updateField} placeholder="e.g. Riya Patnaik" required disabled={!isLoaded || isSaving} />
          </label>
          <label>Email address *
            <input name="email" type="email" value={form.email} onChange={updateField} placeholder="riya@example.com" required disabled={!isLoaded || isSaving} />
          </label>
          <label>Phone number
            <input name="phone" value={form.phone} onChange={updateField} placeholder="Optional" disabled={!isLoaded || isSaving} />
          </label>
          <div className="form-actions">
            <button className="primary" type="submit" disabled={!isLoaded || isSaving}>
              {isSaving ? "Saving…" : editingId ? "Save changes to database" : "Add contact to database"}
            </button>
            {editingId && <button className="secondary" type="button" onClick={cancelEdit}>Cancel</button>}
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
                  <button className="text-button" onClick={() => editContact(contact)} disabled={!isLoaded || isSaving}>Edit</button>
                  <button className="text-button danger" onClick={() => deleteContact(contact.id)} disabled={!isLoaded || isSaving}>Delete</button>
                </td>
              </tr>)}
              {filtered.length === 0 && <tr><td colSpan="4" className="empty">No contacts found.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
      <footer>Stored locally in this browser · Import an Excel workbook once if needed</footer>
    </main>
  );
}
