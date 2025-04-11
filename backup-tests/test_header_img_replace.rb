require 'bundler/setup'
require 'docx'
require 'tempfile'

# Caminho para o seu documento DOCX (modelo)
DOCX_PATH = 'PI.docx'

# Caminho para as imagens que você quer adicionar
HEADER_IMAGE_PATH = 'header.png'
FOOTER_IMAGE_PATH = 'footer.png'

# Função para adicionar uma imagem ao cabeçalho
def replace_header_image(doc, image_path)
  # Lê a imagem
  image_data = File.binread(image_path)
  
  # Nome da imagem dentro do pacote DOCX
  image_id = "header_image_#{Time.now.to_i}"
  image_rel_id = "rId#{Random.rand(1000)}"
  image_filename = "word/media/#{image_id}.png"
  
  # Adiciona a imagem ao pacote
  doc.replace_entry(image_filename, image_data)
  
  # Acessa os cabeçalhos existentes
  if doc.headers && !doc.headers.empty?
    # Obtém a chave do primeiro cabeçalho
    header_key = doc.headers.keys.first
    header = doc.headers[header_key]
    
    # Atualiza o arquivo de relacionamentos para o cabeçalho
    header_rels_path = "word/_rels/header1.xml.rels"
    
    # Se o arquivo de relacionamentos já existe, lê-o e adiciona a nova relação
    if doc.zip.find_entry(header_rels_path)
      rels_content = doc.zip.read(header_rels_path)
      rels_xml = Nokogiri::XML(rels_content)
      
      # Cria um novo nó de relacionamento para a imagem
      rels_root = rels_xml.at_xpath("//xmlns:Relationships", {"xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships"})
      rel_node = Nokogiri::XML::Node.new("Relationship", rels_xml)
      rel_node['Id'] = image_rel_id
      rel_node['Type'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      rel_node['Target'] = "../media/#{image_id}.png"
      rels_root.add_child(rel_node)
      
      # Atualiza o arquivo de relacionamentos
      doc.replace_entry(header_rels_path, rels_xml.to_xml)
    else
      # Cria um novo arquivo de relacionamentos
      rels_xml = Nokogiri::XML::Builder.new(encoding: 'UTF-8') do |xml|
        xml.Relationships(xmlns: "http://schemas.openxmlformats.org/package/2006/relationships") {
          xml.Relationship(Id: image_rel_id, 
                          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                          Target: "../media/#{image_id}.png")
        }
      end
      doc.replace_entry(header_rels_path, rels_xml.to_xml)
    end
    
    # Agora modifica o XML do cabeçalho para usar a imagem
    if header.is_a?(Nokogiri::XML::Document)
      # Substitui o conteúdo atual do cabeçalho por um parágrafo com a imagem
      # Primeiro, remova qualquer imagem existente
      header.xpath('//w:drawing').each(&:remove)
      
      # Criamos um novo parágrafo
      paragraph = header.at_xpath('//w:p') || Nokogiri::XML::Node.new('w:p', header)
      
      # XML para incluir uma imagem em um documento Word
      drawing_xml = <<~XML
        <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="5000000" cy="1000000"/>
            <wp:docPr id="1" name="Picture 1"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="Image"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="#{image_rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="5000000" cy="1000000"/>
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
      XML
      
      # Adicione o nó de desenho ao parágrafo
      fragment = Nokogiri::XML.fragment(drawing_xml)
      paragraph.add_child(fragment)
      
      # Substitui o XML do cabeçalho
      doc.replace_entry("word/header1.xml", header.to_xml)
      
      puts "Imagem adicionada ao cabeçalho com sucesso."
    else
      puts "O cabeçalho não é do tipo esperado."
    end
  else
    puts "Nenhum cabeçalho encontrado no documento."
  end
end

# Função para adicionar uma imagem ao rodapé
def replace_footer_image(doc, image_path)
  # Lê a imagem
  image_data = File.binread(image_path)
  
  # Nome da imagem dentro do pacote DOCX
  image_id = "footer_image_#{Time.now.to_i}"
  image_rel_id = "rId#{Random.rand(1000)}"
  image_filename = "word/media/#{image_id}.png"
  
  # Adiciona a imagem ao pacote
  doc.replace_entry(image_filename, image_data)
  
  # Acessa os rodapés existentes
  if doc.footers && !doc.footers.empty?
    # Obtém a chave do primeiro rodapé
    footer_key = doc.footers.keys.first
    footer = doc.footers[footer_key]
    
    # Atualiza o arquivo de relacionamentos para o rodapé
    footer_rels_path = "word/_rels/footer1.xml.rels"
    
    # Se o arquivo de relacionamentos já existe, lê-o e adiciona a nova relação
    if doc.zip.find_entry(footer_rels_path)
      rels_content = doc.zip.read(footer_rels_path)
      rels_xml = Nokogiri::XML(rels_content)
      
      # Cria um novo nó de relacionamento para a imagem
      rels_root = rels_xml.at_xpath("//xmlns:Relationships", {"xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships"})
      rel_node = Nokogiri::XML::Node.new("Relationship", rels_xml)
      rel_node['Id'] = image_rel_id
      rel_node['Type'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      rel_node['Target'] = "../media/#{image_id}.png"
      rels_root.add_child(rel_node)
      
      # Atualiza o arquivo de relacionamentos
      doc.replace_entry(footer_rels_path, rels_xml.to_xml)
    else
      # Cria um novo arquivo de relacionamentos
      rels_xml = Nokogiri::XML::Builder.new(encoding: 'UTF-8') do |xml|
        xml.Relationships(xmlns: "http://schemas.openxmlformats.org/package/2006/relationships") {
          xml.Relationship(Id: image_rel_id, 
                          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                          Target: "../media/#{image_id}.png")
        }
      end
      doc.replace_entry(footer_rels_path, rels_xml.to_xml)
    end
    
    # Agora modifica o XML do rodapé para usar a imagem
    if footer.is_a?(Nokogiri::XML::Document)
      # Substitui o conteúdo atual do rodapé por um parágrafo com a imagem
      # Primeiro, remova qualquer imagem existente
      footer.xpath('//w:drawing').each(&:remove)
      
      # Criamos um novo parágrafo
      paragraph = footer.at_xpath('//w:p') || Nokogiri::XML::Node.new('w:p', footer)
      
      # XML para incluir uma imagem em um documento Word
      drawing_xml = <<~XML
        <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="5000000" cy="1000000"/>
            <wp:docPr id="2" name="Picture 2"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="Image"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="#{image_rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="5000000" cy="1000000"/>
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
      XML
      
      # Adicione o nó de desenho ao parágrafo
      fragment = Nokogiri::XML.fragment(drawing_xml)
      paragraph.add_child(fragment)
      
      # Substitui o XML do rodapé
      doc.replace_entry("word/footer1.xml", footer.to_xml)
      
      puts "Imagem adicionada ao rodapé com sucesso."
    else
      puts "O rodapé não é do tipo esperado."
    end
  else
    puts "Nenhum rodapé encontrado no documento."
  end
end

begin
  # Verificar se os arquivos existem
  unless File.exist?(DOCX_PATH)
    puts "Erro: Arquivo DOCX #{DOCX_PATH} não encontrado."
    exit 1
  end
  
  unless File.exist?(HEADER_IMAGE_PATH)
    puts "Erro: Arquivo de imagem para o cabeçalho #{HEADER_IMAGE_PATH} não encontrado."
    exit 1
  end
  
  unless File.exist?(FOOTER_IMAGE_PATH)
    puts "Erro: Arquivo de imagem para o rodapé #{FOOTER_IMAGE_PATH} não encontrado."
    exit 1
  end
  
  # Abrir o documento
  doc = Docx::Document.open(DOCX_PATH)
  
  # Verificar se o documento tem cabeçalhos e rodapés
  puts "Documento carregado: #{DOCX_PATH}"
  
  if doc.headers && !doc.headers.empty?
    puts "\nCabeçalhos encontrados:"
    doc.headers.each do |key, header|
      puts "- #{key}: #{header.text}"
    end
  else
    puts "\nNenhum cabeçalho encontrado."
  end
  
  if doc.footers && !doc.footers.empty?
    puts "\nRodapés encontrados:"
    doc.footers.each do |key, footer|
      puts "- #{key}: #{footer.text}"
    end
  else
    puts "\nNenhum rodapé encontrado."
  end
  
  # Substituir as imagens
  replace_header_image(doc, HEADER_IMAGE_PATH)
  replace_footer_image(doc, FOOTER_IMAGE_PATH)
  
  # Salvar o documento modificado
  new_file_path = "customizado_#{File.basename(DOCX_PATH)}"
  doc.save(new_file_path)
  puts "\nDocumento salvo como: #{new_file_path}"
  
rescue StandardError => e
  puts "Erro: #{e.message}"
  puts e.backtrace
end