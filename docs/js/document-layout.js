/** Aturan struktur dokumen tanpa ketergantungan DOM; tidak mengarang isi peraturan. */
const clean = value => String(value || '').replace(/\s+/g, ' ').trim();

export const startsJointApproval = text => /^Dengan\s+Persetujuan\s+Bersama\b/i.test(clean(text));
export const isDecision = text => /^MEMUTUSKAN\s*:/i.test(clean(text));

export function documentMetadata(rows) {
    const lines = [];
    // Hanya pembukaan: nomor dasar hukum tidak boleh menjadi nomor peraturan ini.
    for (const row of rows) {
        if (row.prefix || row.label?.toLowerCase() === 'lampiran') break;
        for (const block of row.ps) {
            const text = clean(block.text);
            if (/^(?:Menimbang|Mengingat|MEMUTUSKAN|BAB\s|Pasal\s)/i.test(text)) break;
            if (text) lines.push(text);
        }
    }
    const heading = lines.join(' ');
    const regulation = heading.match(/\bPERATURAN\s+(?:DAERAH\s+(?:KABUPATEN|KOTA|PROVINSI)|BUPATI|WALI\s*KOTA|GUBERNUR)\s+(.+?)(?=\s+(?:NOMOR|TENTANG|DENGAN RAHMAT)|$)/i);
    const type = regulation ? clean(regulation[0]).toUpperCase() : 'PERATURAN …';
    const numberLine = lines.find(text => /^NOMOR\b/i.test(text));
    const number = numberLine?.match(/^NOMOR\s+(.+?)\s+TAHUN\s+(.+?)(?=\s+TENTANG|$)/i);
    const topic = heading.match(/\bTENTANG\s+(.+?)(?=\s+DENGAN\s+RAHMAT|\s+(?:BUPATI|GUBERNUR|WALI\s*KOTA)\s+[^,]+,|$)/i);
    return {
        regulation: type,
        number: number ? clean(number[1]) : '…',
        year: number ? clean(number[2]) : '…',
        title: topic ? clean(topic[1]).toUpperCase() : '…',
    };
}

export function annexHeadingLines(metadata) {
    return ['LAMPIRAN', metadata.regulation,
        'NOMOR ' + metadata.number + ' TAHUN ' + metadata.year,
        'TENTANG', metadata.title
    ];
}

export function isSourceAnnexHeading(text, metadata) {
    const value = clean(text);
    return /^(?:LAMPIRAN\b|PERATURAN\s+(?:DAERAH|BUPATI|GUBERNUR|WALI)|NOMOR\b|TAHUN\b|TENTANG$)/i.test(value) ||
        value.toUpperCase() === metadata.title;
}
