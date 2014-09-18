module DocxGenerator
  module DSL
    # Represent the docx document.
    class Document
      # Filename of the document (without the docx extension).
      attr_reader :filename
    
      # Create a new docx document.
      # @param filename [String] The filename of the docx file, without the docx extension.
      def initialize(filename, &block)
        @filename = filename + ".docx"
        @objects = [] # It contains all the DSL elements
        yield self if block
      end
      
      # Save the docx document to the target location.
      def save
        generate_archive(generate_content_types, generate_rels, generate_document)
      end

      # Add a new paragraph to the document.
      # @param options [Hash] Formatting options for the paragraph. See the full list in DocxGenerator::DSL::Paragraph.
      def paragraph(options = {}, &block)
        par = DocxGenerator::DSL::Paragraph.new(options)
        yield par if block
        @objects << par
      end

      # Add other objects to the document.
      # @param objects [Object] Objects (like paragraphs).
      # @return [DocxGenerator::DSL::Document] The document object.
      def add(*objects)
        objects.each do |object|
          @objects << object
        end
        self
      end

      def add_table(claim_details, total)    
        "<w:tbl>
          <w:tblPr>
            <w:jc w:val=\"left\"/>
            <w:tblInd w:type=\"dxa\" w:w=\"-15\"/>
            <w:tblBorders>
              <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
              <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
              <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
              <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
              <w:right w:val=\"nil\"/>
              <w:insideV w:val=\"nil\"/>
            </w:tblBorders>
            <w:tblCellMar>
              <w:top w:type=\"dxa\" w:w=\"0\"/>
              <w:left w:type=\"dxa\" w:w=\"93\"/>
              <w:bottom w:type=\"dxa\" w:w=\"0\"/>
              <w:right w:type=\"dxa\" w:w=\"108\"/>
            </w:tblCellMar>
          </w:tblPr>
          <w:tblGrid>
            <w:gridCol w:w=\"523\"/>
            <w:gridCol w:w=\"1809\"/>
            <w:gridCol w:w=\"3262\"/>
            <w:gridCol w:w=\"1438\"/>
            <w:gridCol w:w=\"4\"/>
            <w:gridCol w:w=\"1970\"/>
          </w:tblGrid>
          <w:tr>
            <w:trPr>
              <w:cantSplit w:val=\"false\"/>
            </w:trPr>
            <w:tc>
              <w:tcPr>
                <w:tcW w:type=\"dxa\" w:w=\"523\"/>
                <w:tcBorders>
                  <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:right w:val=\"nil\"/>
                  <w:insideV w:val=\"nil\"/>
                </w:tcBorders>
                <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                <w:tcMar>
                  <w:left w:type=\"dxa\" w:w=\"93\"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val=\"Normal\"/>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                  <w:t>No</w:t>
                </w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:tcPr>
                <w:tcW w:type=\"dxa\" w:w=\"1809\"/>
                <w:tcBorders>
                  <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:right w:val=\"nil\"/>
                  <w:insideV w:val=\"nil\"/>
                </w:tcBorders>
                <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                <w:tcMar>
                  <w:left w:type=\"dxa\" w:w=\"93\"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val=\"Normal\"/>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                  <w:t>Kode Barang</w:t>
                </w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:tcPr>
                <w:tcW w:type=\"dxa\" w:w=\"3262\"/>
                <w:tcBorders>
                  <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:right w:val=\"nil\"/>
                  <w:insideV w:val=\"nil\"/>
                </w:tcBorders>
                <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                <w:tcMar>
                  <w:left w:type=\"dxa\" w:w=\"93\"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val=\"Normal\"/>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                  <w:t>Nama Barang</w:t>
                </w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:tcPr>
                <w:tcW w:type=\"dxa\" w:w=\"1438\"/>
                <w:tcBorders>
                  <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:right w:val=\"nil\"/>
                  <w:insideV w:val=\"nil\"/>
                </w:tcBorders>
                <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                <w:tcMar>
                  <w:left w:type=\"dxa\" w:w=\"93\"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val=\"Normal\"/>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                  <w:t>Quantity</w:t>
                </w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:tcPr>
                <w:tcW w:type=\"dxa\" w:w=\"1974\"/>
                <w:gridSpan w:val=\"2\"/>
                <w:tcBorders>
                  <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:right w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideV w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                </w:tcBorders>
                <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                <w:tcMar>
                  <w:left w:type=\"dxa\" w:w=\"93\"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val=\"Normal\"/>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Bookman Old Style\" w:cs=\"Bookman Old Style\" w:hAnsi=\"Bookman Old Style\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:sz w:val=\"22\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                  <w:t>Biaya</w:t>
                </w:r>
              </w:p>
            </w:tc>
          </w:tr>
          <w:tr>
              <w:trPr>
                <w:cantSplit w:val=\"false\"/>
              </w:trPr>" + generate_table_content(claim_details, total) +
            "</w:tr>
          <w:tr>
            <w:trPr>
              <w:cantSplit w:val=\"false\"/>
            </w:trPr>
            <w:tc>
              <w:tcPr>
                <w:tcW w:type=\"dxa\" w:w=\"7036\"/>
                <w:gridSpan w:val=\"5\"/>
                <w:tcBorders>
                  <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:right w:val=\"nil\"/>
                  <w:insideV w:val=\"nil\"/>
                </w:tcBorders>
                <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                <w:tcMar>
                  <w:left w:type=\"dxa\" w:w=\"93\"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val=\"Normal\"/>
                  <w:jc w:val=\"center\"/>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:color w:val=\"000000\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:color w:val=\"000000\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                  <w:t>Grand Total</w:t>
                </w:r>
              </w:p>
            </w:tc>
            <w:tc>
              <w:tcPr>
                <w:tcW w:type=\"dxa\" w:w=\"1970\"/>
                <w:tcBorders>
                  <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:right w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                  <w:insideV w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                </w:tcBorders>
                <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                <w:tcMar>
                  <w:left w:type=\"dxa\" w:w=\"93\"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val=\"Normal\"/>
                  <w:jc w:val=\"right\"/>
                  <w:rPr>
                    <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                    <w:b/>
                    <w:color w:val=\"000000\"/>
                  </w:rPr>
                </w:pPr>
                <w:r>
                  <w:rPr>
                    <w:rStyle w:val=\"Emphasis\"/>
                    <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                    <w:b/>
                    <w:i w:val=\"false\"/>
                    <w:color w:val=\"000000\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>
                  <w:t xml:space=\"preserve\">Rp </w:t>
                </w:r>
                <w:r>
                  <w:rPr>
                    <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                    <w:b/>
                    <w:color w:val=\"000000\"/>
                    <w:szCs w:val=\"22\"/>
                  </w:rPr>                  
                </w:r>
                <w:r>
                  <w:rPr>
                    <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                    <w:b/>
                    <w:color w:val=\"000000\"/>
                  </w:rPr>
                  <w:t>#{total}</w:t>
                </w:r>
              </w:p>
            </w:tc>
          </w:tr>  
        </w:tbl>"
      end

      def generate_table_content(claim_details, total)
        claim_details.each_with_index do |detail, no| 
          return "<w:tc>
            <w:tcPr>
              <w:tcW w:type=\"dxa\" w:w=\"523\"/>
              <w:tcBorders>
                <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:right w:val=\"nil\"/>
                <w:insideV w:val=\"nil\"/>
              </w:tcBorders>
              <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
              <w:tcMar>
                <w:left w:type=\"dxa\" w:w=\"93\"/>
              </w:tcMar>
            </w:tcPr>
            <w:p>
              <w:pPr>
                <w:pStyle w:val=\"Normal\"/>
                <w:rPr>
                  <w:rStyle w:val=\"Emphasis\"/>
                  <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                  <w:i w:val=\"false\"/>
                  <w:color w:val=\"000000\"/>
                  <w:szCs w:val=\"22\"/>
                </w:rPr>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:rStyle w:val=\"Emphasis\"/>
                  <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                  <w:i w:val=\"false\"/>
                  <w:color w:val=\"000000\"/>
                  <w:szCs w:val=\"22\"/>
                </w:rPr>
                <w:t>#{(no+1).to_s}</w:t>
              </w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:type=\"dxa\" w:w=\"1809\"/>
              <w:tcBorders>
                <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                <w:right w:val=\"nil\"/>
                <w:insideV w:val=\"nil\"/>
              </w:tcBorders>
              <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
              <w:tcMar>
                <w:left w:type=\"dxa\" w:w=\"93\"/>
              </w:tcMar>
            </w:tcPr>
            <w:p>
              <w:pPr>
                <w:pStyle w:val=\"Normal\"/>
                <w:rPr>
                  <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                  <w:color w:val=\"000000\"/>
                  <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                </w:rPr>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                  <w:color w:val=\"000000\"/>
                  <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                </w:rPr>
                <w:t>#{detail.code_barang}</w:t>
              </w:r>
            </w:p>
          </w:tc>
          <w:tc>
                  <w:tcPr>
                    <w:tcW w:type=\"dxa\" w:w=\"3262\"/>
                    <w:tcBorders>
                      <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:right w:val=\"nil\"/>
                      <w:insideV w:val=\"nil\"/>
                    </w:tcBorders>
                    <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                    <w:tcMar>
                      <w:left w:type=\"dxa\" w:w=\"93\"/>
                    </w:tcMar>
                  </w:tcPr>
                  <w:p>
                    <w:pPr>
                      <w:pStyle w:val=\"Normal\"/>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Arial\" w:hAnsi=\"Calibri\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                    </w:pPr>
                    <w:r>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Arial\" w:hAnsi=\"Calibri\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                      <w:t>#{detail.nama_barang}</w:t>
                    </w:r>
                  </w:p>
                </w:tc>
          <w:tc>
                  <w:tcPr>
                    <w:tcW w:type=\"dxa\" w:w=\"1438\"/>
                    <w:tcBorders>
                      <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:right w:val=\"nil\"/>
                      <w:insideV w:val=\"nil\"/>
                    </w:tcBorders>
                    <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                    <w:tcMar>
                      <w:left w:type=\"dxa\" w:w=\"93\"/>
                    </w:tcMar>
                  </w:tcPr>
                  <w:p>
                    <w:pPr>
                      <w:pStyle w:val=\"Normal\"/>
                      <w:jc w:val=\"center\"/>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                        <w:color w:val=\"000000\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                    </w:pPr>
                    <w:r>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                        <w:color w:val=\"000000\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                      <w:t>#{detail.jumlah}</w:t>
                    </w:r>
                  </w:p>
                </w:tc>
          <w:tc>
                  <w:tcPr>
                    <w:tcW w:type=\"dxa\" w:w=\"1974\"/>
                    <w:gridSpan w:val=\"2\"/>
                    <w:tcBorders>
                      <w:top w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:left w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:bottom w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:insideH w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:right w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                      <w:insideV w:color=\"000001\" w:space=\"0\" w:sz=\"4\" w:val=\"single\"/>
                    </w:tcBorders>
                    <w:shd w:fill=\"FFFFFF\" w:val=\"clear\"/>
                    <w:tcMar>
                      <w:left w:type=\"dxa\" w:w=\"93\"/>
                    </w:tcMar>
                  </w:tcPr>
                  <w:p>
                    <w:pPr>
                      <w:pStyle w:val=\"Normal\"/>
                      <w:jc w:val=\"right\"/>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                        <w:color w:val=\"000000\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                    </w:pPr>
                    <w:r>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                        <w:color w:val=\"000000\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                      <w:t>#{detail.biaya}</w:t>
                    </w:r>
                  </w:p>
                  <w:p>
                    <w:pPr>
                      <w:pStyle w:val=\"Normal\"/>
                      <w:jc w:val=\"right\"/>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                        <w:color w:val=\"000000\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                    </w:pPr>
                    <w:r>
                      <w:rPr>
                        <w:rFonts w:ascii=\"Calibri\" w:cs=\"Calibri\" w:hAnsi=\"Calibri\"/>
                        <w:color w:val=\"000000\"/>
                        <w:shd w:fill=\"FFFF00\" w:val=\"clear\"/>
                      </w:rPr>
                      <w:t>(inc. PPn)</w:t>
                    </w:r>
                  </w:p>
                </w:tc>"
        end
      end

      def table(claim_details, total)
        @objects << add_table(claim_details, total)
      end
      
      private
      
        def generate_content_types
          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
          <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
            <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
            <Default Extension=\"xml\" ContentType=\"application/xml\"/>
            <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>
          </Types>"
        end
        
        def generate_rels
          <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
EOF
        end
        
        def generate_document
          content = []
          @objects.each do |object|
            content << object.generate
          end

          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          Word::Document.new({ "xmlns:w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main" },
            [ Word::Body.new(content) ]).to_s
        end

        def generate_archive(content_types, rels, document)
          File.delete(@filename) if File.exists?(@filename)
          Zip::File.open(@filename, Zip::File::CREATE) do |docx|
            docx.mkdir('_rels')
            docx.mkdir('word')
            docx.get_output_stream('[Content_Types].xml') { |f| f.puts content_types }
            docx.get_output_stream('_rels/.rels') { |f| f.puts rels }
            docx.get_output_stream('word/document.xml') { |f| f.puts document }
          end
        end
    end
  end
end
