(async function () {
  const $ = (id) => document.getElementById(id);
  const log = (msg) => {
    $("status").textContent = String(msg);
  };

  function uniqueFilename(base = "merged_output.docx") {
    return base;
  }

  function getTextFromCell(cellEl) {
    const TEXT_NS =
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let result = "";
    const walker = (
      cellEl && cellEl.ownerDocument ? cellEl.ownerDocument : document
    ).createTreeWalker(cellEl, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const n = walker.currentNode;
      if (n.nodeType === 1) {
        if (n.namespaceURI === TEXT_NS && n.localName === "br") result += "\n";
      } else if (n.nodeType === 3) {
        result += n.nodeValue;
      }
    }
    return result.replace(/\r?\n/g, "\n").trim();
  }

  function removeLeadingId(text) {
    return (text || "").trim().replace(/^[^\s]+\s*/, "");
  }

  async function readDocxArrayBuffer(file) {
    return new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onerror = () => rej(fr.error);
      fr.onload = () => res(fr.result);
      fr.readAsArrayBuffer(file);
    });
  }

  async function extractTableData(file) {
    const ab = await readDocxArrayBuffer(file);
    const zip = await JSZip.loadAsync(ab);
    const docXml = await zip.file("word/document.xml").async("string");
    const parser = new DOMParser();
    const xml = parser.parseFromString(docXml, "application/xml");

    const WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const tables = xml.getElementsByTagNameNS(WNS, "tbl");
    if (!tables || !tables.length)
      throw new Error("No tables found in document.");

    const allTables = [];

    Array.from(tables).forEach((tbl) => {
      const rows = Array.from(tbl.getElementsByTagNameNS(WNS, "tr"));
      const questions = [];
      const ids = [];

      rows.forEach((tr) => {
        const cells = Array.from(tr.getElementsByTagNameNS(WNS, "tc"));
        if (cells.length >= 2) {
          const text = getTextFromCell(cells[1]);
          questions.push(text);

          if (
            !isOptionRow(text) &&
            !/^ANSWER/i.test(text) &&
            text.trim() !== ""
          ) {
            const firstToken = text.trim().split(/\s+/)[0] || "";
            ids.push(firstToken);
          }
        }
      });

      allTables.push({ questions, ids });
    });

    return allTables;
  }

  function isOptionRow(text) {
    return /^[A-Z]\s*[\.\)]/.test((text || "").trim());
  }

  function buildMergedRows(q1, q2, ids, lang1, lang2) {
    if (q1.length !== q2.length)
      throw new Error("The number of lines in the files does not match!");

    const rows = [];
    let qNumber = 1;
    let idIndex = 0;

    for (let i = 0; i < q1.length; i++) {
      const t1 = q1[i] || "";
      const t2 = q2[i] || "";

      if (!isOptionRow(t1) && !/^ANSWER/i.test(t1) && t1.trim() !== "") {
        const currentId = ids[idIndex++] || "";
        const clean1 = removeLeadingId(t1);
        const clean2 = removeLeadingId(t2);
        rows.push({
          left: String(qNumber++),
          middle: `${currentId}  {mlang ${lang1}}${clean1}{mlang}{mlang ${lang2}}${clean2}{mlang}`,
          right: "",
        });
      } else if (isOptionRow(t1)) {
        const prefix = t1.slice(0, 2);
        const val1 = t1.slice(2).trim();
        const val2 = (
          t2.trim().toUpperCase().startsWith(prefix) ? t2.slice(2) : t2
        ).trim();
        rows.push({
          left: "",
          middle: `${prefix} {mlang ${lang1}}${val1}{mlang}{mlang ${lang2}}${val2}{mlang}`,
          right: "",
        });
      } else if (/^ANSWER/i.test(t1)) {
        rows.push({ left: "", middle: t1, right: "" });
      } else {
        rows.push({ left: "", middle: t1, right: "" });
      }
    }
    return rows;
  }

  async function createDocxAndDownload(allTablesRows) {
    const {
      Document,
      Packer,
      Paragraph,
      Table,
      TableRow,
      TableCell,
      WidthType,
      BorderStyle,
      TextRun,
    } = docx;

    const toParagraph = (text) =>
      new Paragraph({
        children: [new TextRun(text)],
        spacing: { before: 0, after: 0, line: 240 },
      });

    const docChildren = [];

    allTablesRows.forEach((rows, idx) => {
      const tableRows = rows.map(
        (r) =>
          new TableRow({
            children: [
              new TableCell({
                children: [toParagraph(r.left || "")],
                width: { size: 720, type: WidthType.DXA },
              }),
              new TableCell({
                children: [toParagraph(r.middle || "")],
                width: { size: 6480, type: WidthType.DXA },
              }),
              new TableCell({
                children: [toParagraph(r.right || "")],
                width: { size: 720, type: WidthType.DXA },
              }),
            ],
          })
      );

      const table = new Table({
        rows: tableRows,
        width: { size: 7920, type: WidthType.DXA },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
          left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
          right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
          insideHorizontal: {
            style: BorderStyle.SINGLE,
            size: 4,
            color: "000000",
          },
          insideVertical: {
            style: BorderStyle.SINGLE,
            size: 4,
            color: "000000",
          },
        },
      });

      docChildren.push(table);

      if (idx < allTablesRows.length - 1) {
        docChildren.push(new Paragraph({ text: "", spacing: { after: 200 } }));
      }
    });

    const doc = new Document({
      sections: [{ properties: {}, children: docChildren }],
    });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, uniqueFilename("merged_output.docx"));
  }

  $("mergeBtn").addEventListener("click", async () => {
    try {
      const f1 = $("file1").files[0];
      const f2 = $("file2").files[0];
      const lang1 = ($("lang1").value || "").trim();
      const lang2 = ($("lang2").value || "").trim();

      if (!f1 || !f2) throw new Error("You must select both files (.docx).");
      if (!lang1 || !lang2)
        throw new Error("You must enter both language codes.");

      log("Reading first document...");
      const tables1 = await extractTableData(f1);

      log("Reading second document...");
      const tables2 = await extractTableData(f2);

      if (tables1.length !== tables2.length) {
        throw new Error("The number of tables in the files does not match!");
      }

      log("Building merged tables...");
      const allTablesRows = tables1.map((t1, idx) =>
        buildMergedRows(
          t1.questions,
          tables2[idx].questions,
          t1.ids,
          lang1,
          lang2
        )
      );

      log("Generating .docx...");
      await createDocxAndDownload(allTablesRows);
      log("Done. File downloaded: merged_output.docx");
    } catch (err) {
      console.error(err);
      log("Error: " + (err && err.message ? err.message : String(err)));
      alert("Error: " + (err && err.message ? err.message : String(err)));
    }
  });

  // --- Self-test ---
  function assertEquals(actual, expected, label) {
    const ok = actual === expected;
    return ok
      ? `✅ ${label}`
      : `❌ ${label} (got: ${JSON.stringify(
          actual
        )}, expected: ${JSON.stringify(expected)})`;
  }

  function runTests() {
    const out = [];
    // removeLeadingId
    out.push(
      assertEquals(
        removeLeadingId("Q1 What is this?"),
        "What is this?",
        "removeLeadingId strips first token"
      )
    );
    out.push(
      assertEquals(
        removeLeadingId("  ABC123  Hello world "),
        "Hello world",
        "removeLeadingId trims"
      )
    );

    // isOptionRow
    out.push(assertEquals(isOptionRow("A. Cat"), true, "isOptionRow A. true"));
    out.push(assertEquals(isOptionRow("Z."), true, "isOptionRow Z. true"));
    out.push(
      assertEquals(isOptionRow("1. Not"), false, "isOptionRow number false")
    );
    out.push(assertEquals(isOptionRow("B) Dog"), true, "isOptionRow B) true"));

    // buildMergedRows – jedan kompletan set (npr. 4 odgovora)
    const q1 = [
      "QID001 What is your favorite animal?",
      "A. Cat",
      "B. Dog",
      "C. Bird",
      "D. Fish",
      "ANSWER: B",
      "",
    ];
    const q2 = [
      "QID001 Koja je vaša omiljena životinja?",
      "A. Mačka",
      "B. Pas",
      "C. Ptica",
      "D. Riba",
      "ODGOVOR: B",
      "",
    ];
    const ids = ["QID001"];
    const rows = buildMergedRows(q1, q2, ids, "en", "uk");
    out.push(assertEquals(rows.length, 7, "rows built for 1 question"));
    out.push(assertEquals(rows[0].left, "1", "question number 1 in left cell"));
    out.push(
      assertEquals(
        rows[1].middle.startsWith(
          "A. {mlang en}Cat{mlang}{mlang uk}Mačka{mlang}"
        ),
        true,
        "option A bilingual merged"
      )
    );
    out.push(
      assertEquals(rows[5].middle, "ANSWER: B", "ANSWER copied from Lang1")
    );

    log(out.join("\n"));
  }

  $("testBtn").addEventListener("click", runTests);
})();
