/**
 * Satu sumber keadaan dokumen. Modul lain mengimpor binding model yang sama.
 */
export let model = null;
export function setModel(value) {
    model = value;
    window.KOMAD_DOCUMENT_OPEN = !!value;
}
