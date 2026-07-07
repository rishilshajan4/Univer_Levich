/**
 * Version store (VH-1) — a tiny, dependency-free, browser-only checkpoint store.
 *
 * A "version" is a document-level checkpoint: the whole manifest + every sheet's
 * snapshot + the title, at a moment, plus metadata (who / when / why). The store
 * is append-only in spirit (the host only ever `putVersion`s new ids, or updates a
 * label on an existing one) and keyed by `documentId` so several documents can
 * share one database.
 *
 * Persistence is IndexedDB (instant local autosave + survives reload). When
 * IndexedDB is unavailable (SSR, private mode, quota errors) it transparently
 * falls back to an in-memory Map so the feature still works for the session.
 *
 * No external deps, no Node APIs, no `@univerjs-pro/*` — Constitution II/IX.
 */
import type { SheetManifestEntry, SingleSheetSnapshot } from "../core/shell-workbook";

/** A full document checkpoint: the manifest + every sheet's snapshot + the title. */
export interface DocumentSnapshot {
  manifest: SheetManifestEntry[];
  sheets: Record<string, SingleSheetSnapshot>;
  title: string;
}

/** Why a checkpoint was cut. `blank`/`import` mark the immutable seq-0 original. */
export type VersionKind = "import" | "blank" | "auto" | "named" | "restore";

/** One stored checkpoint. */
export interface Version {
  id: string;
  /** Monotonic order (0 = the original). */
  seq: number;
  /** Set = a named version (kept forever); undefined = auto/original. */
  label?: string;
  kind: VersionKind;
  /** Author display name (single-user demo = "You"). */
  author: string;
  createdAt: number;
  document: DocumentSnapshot;
}

/** The public store contract the host wires to. */
export interface VersionStore {
  listVersions(documentId: string): Promise<Version[]>;
  putVersion(documentId: string, version: Version): Promise<void>;
  getVersion(documentId: string, versionId: string): Promise<Version | undefined>;
  /** Remove a version (used to prune old auto-checkpoints — see the host's retention policy). */
  deleteVersion(documentId: string, versionId: string): Promise<void>;
}

const DB_NAME = "finsheets-version-history";
const STORE = "versions";
const DB_VERSION = 1;

/** Composite primary key so one DB holds versions for many documents. */
function pk(documentId: string, id: string): string {
  return `${documentId}::${id}`;
}

interface StoredRecord {
  pk: string;
  documentId: string;
  version: Version;
}

/** Promise-wrap an IDBRequest. */
function reqToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/** In-memory fallback (and the backing store when IndexedDB fails at runtime). */
class MemoryVersionStore implements VersionStore {
  private byDoc = new Map<string, Map<string, Version>>();

  private bucket(documentId: string): Map<string, Version> {
    let m = this.byDoc.get(documentId);
    if (!m) { m = new Map(); this.byDoc.set(documentId, m); }
    return m;
  }

  async listVersions(documentId: string): Promise<Version[]> {
    return [...this.bucket(documentId).values()].sort((a, b) => a.seq - b.seq);
  }
  async putVersion(documentId: string, version: Version): Promise<void> {
    this.bucket(documentId).set(version.id, version);
  }
  async getVersion(documentId: string, versionId: string): Promise<Version | undefined> {
    return this.bucket(documentId).get(versionId);
  }
  async deleteVersion(documentId: string, versionId: string): Promise<void> {
    this.bucket(documentId).delete(versionId);
  }
}

/** IndexedDB-backed store; every method degrades to memory on any error. */
class IdbVersionStore implements VersionStore {
  private dbPromise: Promise<IDBDatabase> | null = null;
  private memory = new MemoryVersionStore();

  private open(): Promise<IDBDatabase> {
    if (this.dbPromise) return this.dbPromise;
    this.dbPromise = new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) {
          const os = db.createObjectStore(STORE, { keyPath: "pk" });
          os.createIndex("documentId", "documentId", { unique: false });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    return this.dbPromise;
  }

  async listVersions(documentId: string): Promise<Version[]> {
    try {
      const db = await this.open();
      const tx = db.transaction(STORE, "readonly");
      const index = tx.objectStore(STORE).index("documentId");
      const records = await reqToPromise<StoredRecord[]>(index.getAll(documentId) as IDBRequest<StoredRecord[]>);
      return records.map((r) => r.version).sort((a, b) => a.seq - b.seq);
    } catch {
      return this.memory.listVersions(documentId);
    }
  }

  async putVersion(documentId: string, version: Version): Promise<void> {
    // Mirror into memory too, so a later IndexedDB failure still has the data.
    await this.memory.putVersion(documentId, version);
    try {
      const db = await this.open();
      const tx = db.transaction(STORE, "readwrite");
      const record: StoredRecord = { pk: pk(documentId, version.id), documentId, version };
      await reqToPromise(tx.objectStore(STORE).put(record));
    } catch {
      /* memory already holds it */
    }
  }

  async getVersion(documentId: string, versionId: string): Promise<Version | undefined> {
    try {
      const db = await this.open();
      const tx = db.transaction(STORE, "readonly");
      const record = await reqToPromise<StoredRecord | undefined>(
        tx.objectStore(STORE).get(pk(documentId, versionId)) as IDBRequest<StoredRecord | undefined>,
      );
      return record?.version ?? (await this.memory.getVersion(documentId, versionId));
    } catch {
      return this.memory.getVersion(documentId, versionId);
    }
  }

  async deleteVersion(documentId: string, versionId: string): Promise<void> {
    await this.memory.deleteVersion(documentId, versionId);
    try {
      const db = await this.open();
      const tx = db.transaction(STORE, "readwrite");
      await reqToPromise(tx.objectStore(STORE).delete(pk(documentId, versionId)));
    } catch {
      /* memory already dropped it */
    }
  }
}

/** Build a version store: IndexedDB when available, else in-memory. */
export function createVersionStore(): VersionStore {
  if (typeof indexedDB !== "undefined") {
    try {
      return new IdbVersionStore();
    } catch {
      /* fall through to memory */
    }
  }
  return new MemoryVersionStore();
}
