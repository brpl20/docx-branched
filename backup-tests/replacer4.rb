require 'bundler/setup'
require 'docx'
require 'tempfile'
require 'fileutils'

# Caminho para o seu documento DOCX (modelo)
DOCX_PATH = 'seu_modelo.docx'

# Caminho para as imagens que você quer adicionar
HEADER_IMAGE_PATH = 'imagem_header.png'
FOOTER_IMAGE_PATH = 'imagem_footer.png'

def replace_header_footer(doc, image_path, is_header=true)
  # Verificar se a imagem existe
  unless File.exist?(image_path)
    puts "Arquivo de imagem #{image_path} não encontrado."
    return false
  end
  
  # Determinar nomes de arquivos
  element_type = is_header ? "header" : "footer"
  xml_file1 = "word/#{element_type}1.xml"
  xml_file2 = "word/#{element_type}2.xml"
  rels_file1 = "word/_rels/#{element_type}1.xml.rels"
  rels_file2 = "word/_rels/#{element_type}2.xml.rels"
  
  # Extrair e copiar a imagem
  image_data = File.binread(image_path)
  image_basename = File.basename(image_path)
  media_path = "word/media/#{image_basename}"
  
  # Adicionar a imagem ao pacote
  doc.replace_entry(media_path, image_data)
  puts "Imagem adicionada ao pacote: #{media_path}"
  
  # Criar arquivos de relacionamento
  rels_content = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" 
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
                    Target="../media/#{image_basename}"/>
    </Relationships>
  XML
  
  # Substituir os arquivos de relacionamento
  doc.replace_entry(rels_file1, rels_content)
  puts "Arquivo de relacionamento atualizado: #{rels_file1}"
  
  if doc.zip.find_entry(rels_file2)
    doc.replace_entry(rels_file2, rels_content)
    puts "Arquivo de relacionamento atualizado: #{rels_file2}"
  end
  
  # Criar conteúdo XML completo para o cabeçalho/rodapé
  xml_content = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:#{element_type} 
      xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" 
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" 
      xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" 
      xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" 
      xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" 
      xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" 
      xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" 
      xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" 
      xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" 
      xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" 
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
      xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" 
      xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" 
      xmlns:o="urn:schemas-microsoft-com:office:office" 
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
      xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" 
      xmlns:v="urn:schemas-microsoft-com:vml" 
      xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" 
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
      xmlns:w10="urn:schemas-microsoft-com:office:word" 
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
      xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" 
      xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" 
      xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" 
      xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" 
      xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" 
      xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" 
      xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" 
      mc:Ignorable="w14 w15 w16se w16cid wp14">
      <w:p>
        <w:pPr>
          <w:pStyle w:val="#{is_header ? 'Cabealho' : 'Rodap'}"/>
          <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:noProof/>
            <w:lang w:eastAsia="pt-BR"/>
          </w:rPr>
          <w:drawing>
            <wp:inline distT="0" distB="0" distL="0" distR="0">
              <wp:extent cx="3022600" cy="669290"/>
              <wp:effectExtent l="0" t="0" r="6350" b="0"/>
              <wp:docPr id="1" name="Imagem 1"/>
              <wp:cNvGraphicFramePr>
                <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
              </wp:cNvGraphicFramePr>
              <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:nvPicPr>
                      <pic:cNvPr id="1" name="Imagem 1"/>
                      <pic:cNvPicPr>
                        <a:picLocks noChangeAspect="1"/>
                      </pic:cNvPicPr>
                    </pic:nvPicPr>
                    <pic:blipFill>
                      <a:blip r:embed="rId1" cstate="print">
                        <a:extLst>
                          <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                            <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                          </a:ext>
                        </a:extLst>
                      </a:blip>
                      <a:stretch>
                        <a:fillRect/>
                      </a:stretch>
                    </pic:blipFill>
                    <pic:spPr>
                      <a:xfrm>
                        <a:off x="0" y="0"/>
                        <a:ext cx="3265882" cy="723135"/>
                      </a:xfrm>
                      <a:prstGeom prst="rect">
                        <a:avLst/>
                      </a:prstGeom>
                    </pic:spPr>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing>
        </w:r>
      </w:p>
    </w:#{element_type}>
  XML
  
  # Substituir completamente os arquivos XML
  doc.replace_entry(xml_file1, xml_content)
  puts "Arquivo XML atualizado: #{xml_file1}"
  
  if doc.zip.find_entry(xml_file2)
    doc.replace_entry(xml_file2, xml_content)
    puts "Arquivo XML atualizado: #{xml_file2}"
  end
  
  return true
end

# Método para garantir que document.xml.rels tenha as referências corretas
def ensure_header_footer_references(doc)
  # Verificar se o arquivo existe
  if doc.zip.find_entry("word/_rels/document.xml.rels")
    # Ler o arquivo existente
    rels_content = doc.zip.read("word/_rels/document.xml.rels")
    rels_xml = Nokogiri::XML(rels_content)
    
    # Verificar e adicionar referências aos cabeçalhos e rodapés
    relationships = rels_xml.xpath("//xmlns:Relationships", {"xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships"}).first
    
    # Verificar e adicionar cabeçalho padrão
    header_default = rels_xml.xpath("//xmlns:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'][@Target='header1.xml']", {"xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships"})
    if header_default.empty?
      header_node = Nokogiri::XML::Node.new("Relationship", rels_xml)
      header_node['Id'] = "rId" + (relationships.children.length + 1).to_s
      header_node['Type'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
      header_node['Target'] = "header1.xml"
      relationships.add_child(header_node)
      puts "Adicionada referência para o cabeçalho padrão"
    end
    
    # Verificar e adicionar rodapé padrão
    footer_default = rels_xml.xpath("//xmlns:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'][@Target='footer1.xml']", {"xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships"})
    if footer_default.empty?
      footer_node = Nokogiri::XML::Node.new("Relationship", rels_xml)
      footer_node['Id'] = "rId" + (relationships.children.length + 1).to_s
      footer_node['Type'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
      footer_node['Target'] = "footer1.xml"
      relationships.add_child(footer_node)
      puts "Adicionada referência para o rodapé padrão"
    end
    
    # Salvar o arquivo atualizado
    doc.replace_entry("word/_rels/document.xml.rels", rels_xml.to_xml)
    puts "Arquivo document.xml.rels atualizado"
  end
end

begin
  # Verificar se o arquivo DOCX existe
  unless File.exist?(DOCX_PATH)
    puts "Erro: Arquivo DOCX #{DOCX_PATH} não encontrado."
    exit 1
  end
  
  # Verificar se as imagens existem
  unless File.exist?(HEADER_IMAGE_PATH)
    puts "Erro: Arquivo de imagem para cabeçalho #{HEADER_IMAGE_PATH} não encontrado."
    exit 1
  end
  
  unless File.exist?(FOOTER_IMAGE_PATH)
    puts "Erro: Arquivo de imagem para rodapé #{FOOTER_IMAGE_PATH} não encontrado."
    exit 1
  end
  
  # Criar uma cópia do arquivo original antes de modificá-lo
  temp_file = "temp_#{File.basename(DOCX_PATH)}"
  FileUtils.cp(DOCX_PATH, temp_file)
  
  # Abrir o documento
  doc = Docx::Document.open(temp_file)
  
  # Substituir cabeçalhos e rodapés
  replace_header_footer(doc, HEADER_IMAGE_PATH, true) # true para cabeçalho
  replace_header_footer(doc, FOOTER_IMAGE_PATH, false) # false para rodapé
  
  # Garantir referências nos arquivos
  ensure_header_footer_references(doc)
  
  # Salvar o documento modificado
  new_file_path = "modificado_#{File.basename(DOCX_PATH)}"
  doc.save(new_file_path)
  puts "\nDocumento salvo como: #{new_file_path}"
  
  # Limpar arquivo temporário
  FileUtils.rm(temp_file)
  
rescue StandardError => e
  puts "Erro: #{e.message}"
  puts e.backtrace
end