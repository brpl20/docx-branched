require 'bundler/setup'
require 'docx'
require 'tempfile'

# Caminho para o seu documento DOCX (modelo)
DOCX_PATH = 'seu_modelo.docx'

# Caminho para as imagens que você quer adicionar
HEADER_IMAGE_PATH = 'imagem_header.png'
FOOTER_IMAGE_PATH = 'imagem_footer.png'

def add_image_to_document(doc, image_path, target_part_name, rel_id = "rId1")
  # Verificar se a imagem existe
  unless File.exist?(image_path)
    puts "Arquivo de imagem #{image_path} não encontrado."
    return false
  end
  
  # Lê a imagem
  image_data = File.binread(image_path)
  
  # Obter nome do arquivo de imagem e caminho no pacote
  image_filename = File.basename(image_path)
  target_filename = "word/media/#{image_filename}"
  
  # Adicionar a imagem ao pacote DOCX
  doc.replace_entry(target_filename, image_data)
  puts "Imagem adicionada ao pacote em: #{target_filename}"
  
  # Determinar o nome do arquivo .rels
  rels_path = "word/_rels/#{File.basename(target_part_name)}.rels"
  
  # Criar conteúdo XML para o arquivo .rels
  # IMPORTANTE: Note o caminho correto usando ../media/
  rels_xml = <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="#{rel_id}" 
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
                    Target="../media/#{image_filename}"/>
    </Relationships>
  XML
  
  # Adicionar o arquivo .rels ao pacote
  doc.replace_entry(rels_path, rels_xml)
  puts "Arquivo de relacionamentos atualizado: #{rels_path}"
  
  # Agora, modifica o XML do cabeçalho/rodapé existente para usar a referência correta
  if doc.zip.find_entry(target_part_name)
    # Ler o arquivo existente
    part_content = doc.zip.read(target_part_name)
    xml_doc = Nokogiri::XML(part_content)
    
    # Encontrar o elemento blip e atualizar o ID de relacionamento
    blips = xml_doc.xpath('//a:blip', {'a' => 'http://schemas.openxmlformats.org/drawingml/2006/main'})
    if blips.any?
      blips.each do |blip|
        blip['r:embed'] = rel_id
      end
      
      # Salvar o arquivo modificado de volta ao pacote
      doc.replace_entry(target_part_name, xml_doc.to_xml)
      puts "Arquivo XML atualizado: #{target_part_name}"
      return true
    else
      puts "Aviso: Não foi possível encontrar o elemento <a:blip> no arquivo #{target_part_name}"
    end
  else
    puts "Erro: O arquivo #{target_part_name} não existe no pacote DOCX"
  end
  
  return false
end

begin
  # Verificar se o arquivo DOCX existe
  unless File.exist?(DOCX_PATH)
    puts "Erro: Arquivo DOCX #{DOCX_PATH} não encontrado."
    exit 1
  end
  
  # Abrir o documento
  doc = Docx::Document.open(DOCX_PATH)
  
  # Adicionar imagem aos cabeçalhos
  add_image_to_document(doc, HEADER_IMAGE_PATH, "word/header1.xml")
  add_image_to_document(doc, HEADER_IMAGE_PATH, "word/header2.xml")
  
  # Adicionar imagem aos rodapés
  add_image_to_document(doc, FOOTER_IMAGE_PATH, "word/footer1.xml")
  add_image_to_document(doc, FOOTER_IMAGE_PATH, "word/footer2.xml")
  
  # Salvar o documento modificado
  new_file_path = "corrigido_#{File.basename(DOCX_PATH)}"
  doc.save(new_file_path)
  puts "\nDocumento salvo como: #{new_file_path}"
  
rescue StandardError => e
  puts "Erro: #{e.message}"
  puts e.backtrace
end