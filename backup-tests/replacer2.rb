require 'bundler/setup'
require 'docx'
require 'tempfile'

# Caminho para o seu documento DOCX (modelo)
DOCX_PATH = 'PI.docx'

# Caminho para as imagens que você quer adicionar
HEADER_IMAGE_PATH = 'header-cinza.png'
FOOTER_IMAGE_PATH = 'footer-cinza.png'

def add_image_to_document(doc, image_path, target_file, is_header = true)
  # Verificar se a imagem existe
  unless File.exist?(image_path)
    puts "Arquivo de imagem #{image_path} não encontrado."
    return false
  end
  
  # Lê a imagem
  image_data = File.binread(image_path)
  
  # Gerar um ID único para a imagem
  image_filename = File.basename(image_path)
  timestamp = Time.now.to_i
  image_id = "image_#{timestamp}"
  media_path = "word/media/#{image_id}_#{image_filename}"
  
  # Adicionar a imagem ao pacote DOCX
  doc.replace_entry(media_path, image_data)
  puts "Imagem adicionada ao pacote em: #{media_path}"
  
  # Criar ou atualizar o arquivo .rels
  type = is_header ? "header" : "footer"
  number = "1" # Geralmente header1.xml ou footer1.xml
  rels_path = "word/_rels/#{type}#{number}.xml.rels"
  rel_id = "rId100" # Usar um ID provavelmente não utilizado
  
  # Criar conteúdo XML para o arquivo .rels
  rels_xml = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="#{rel_id}" 
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
                    Target="../media/#{image_id}_#{image_filename}"/>
    </Relationships>
  XML
  
  # Adicionar o arquivo .rels ao pacote
  doc.replace_entry(rels_path, rels_xml)
  puts "Arquivo de relacionamentos atualizado: #{rels_path}"
  
  # Agora criar o conteúdo XML para o cabeçalho/rodapé
  content_xml = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:#{type} 
      xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" 
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" 
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
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
      xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" 
      xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" 
      xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" 
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" 
      mc:Ignorable="w14 w15 wp14">
      <w:p>
        <w:pPr>
          <w:pStyle w:val="#{is_header ? 'Cabealho' : 'Rodap'}"/>
          <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:noProof/>
          </w:rPr>
          <w:drawing>
            <wp:inline distT="0" distB="0" distL="0" distR="0">
              <wp:extent cx="5000000" cy="1500000"/>
              <wp:effectExtent l="0" t="0" r="0" b="0"/>
              <wp:docPr id="1" name="Imagem"/>
              <wp:cNvGraphicFramePr>
                <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
              </wp:cNvGraphicFramePr>
              <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:nvPicPr>
                      <pic:cNvPr id="0" name="Image"/>
                      <pic:cNvPicPr/>
                    </pic:nvPicPr>
                    <pic:blipFill>
                      <a:blip r:embed="#{rel_id}"/>
                      <a:stretch>
                        <a:fillRect/>
                      </a:stretch>
                    </pic:blipFill>
                    <pic:spPr>
                      <a:xfrm>
                        <a:off x="0" y="0"/>
                        <a:ext cx="5000000" cy="1500000"/>
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
    </w:#{type}>
  XML
  
  # Adicionar o conteúdo XML ao pacote
  doc.replace_entry(target_file, content_xml)
  puts "#{is_header ? 'Cabeçalho' : 'Rodapé'} atualizado: #{target_file}"
  
  return true
end

begin
  # Verificar se o arquivo DOCX existe
  unless File.exist?(DOCX_PATH)
    puts "Erro: Arquivo DOCX #{DOCX_PATH} não encontrado."
    exit 1
  end
  
  # Abrir o documento
  doc = Docx::Document.open(DOCX_PATH)
  
  # Adicionar imagem ao cabeçalho
  add_image_to_document(doc, HEADER_IMAGE_PATH, "word/header1.xml", true)
  
  # Adicionar imagem ao rodapé
  add_image_to_document(doc, FOOTER_IMAGE_PATH, "word/footer1.xml", false)
  
  # Garantir que o documento tenha referências aos cabeçalhos e rodapés
  # Isso é importante se você estiver criando um documento do zero ou substituindo completamente
  # Pode ser necessário atualizar document.xml.rels e settings.xml para referenciar os cabeçalhos/rodapés
  
  # Salvar o documento modificado
  new_file_path = "customizado_#{File.basename(DOCX_PATH)}"
  doc.save(new_file_path)
  puts "\nDocumento salvo como: #{new_file_path}"
  
rescue StandardError => e
  puts "Erro: #{e.message}"
  puts e.backtrace
end