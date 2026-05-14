// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class RaTests
    {
        [Theory]
        [InlineData("RA001-Tracked-Revisions-01.docx")]
        [InlineData("RA001-Tracked-Revisions-02.docx")]

        public void RA001(string name)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

            WmlDocument notAccepted = new WmlDocument(sourceDocx.FullName);
            WmlDocument afterAccepting = RevisionAccepter.AcceptRevisions(notAccepted);
            var processedDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx")));
            afterAccepting.SaveAs(processedDestDocx.FullName);
        }

        [Fact]
        public void AcceptRevisionsHandlesAbsentOptionalParts()
        {
            using (var stream = new MemoryStream())
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainDocumentPart = doc.AddMainDocumentPart();
                mainDocumentPart.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text("Hello")))));
                mainDocumentPart.Document.Save();

                RevisionAccepter.AcceptRevisions(doc);
            }
        }

        [Fact]
        public void AcceptRevisionsRejectsDocumentWithoutMainDocumentPart()
        {
            using (var stream = new MemoryStream())
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var exception = Assert.Throws<InvalidOperationException>(() => RevisionAccepter.AcceptRevisions(doc));

                Assert.Equal("The WordprocessingDocument does not contain a MainDocumentPart.", exception.Message);
            }
        }

    }
}

#endif
