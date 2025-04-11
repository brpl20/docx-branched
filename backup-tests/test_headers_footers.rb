require 'bundler/setup'
require 'docx'
require 'tempfile'

# Caminho para um arquivo DOCX de exemplo que tenha cabeçalho e rodapé
DOCX_PATH = 'FDP.docx'

# Função para mostrar informações sobre cabeçalhos e rodapés
def show_doc_info(doc)
  puts "Documento carregado: #{DOCX_PATH}"
  
  # Verificar se há cabeçalhos
  if doc.headers && !doc.headers.empty?
    puts "\nCabeçalhos encontrados:"
    doc.headers.each do |key, header|
      puts "- #{key}: #{header.text}"
    end
  else
    puts "\nNenhum cabeçalho encontrado."
  end
  
  # Verificar se há rodapés
  if doc.footers && !doc.footers.empty?
    puts "\nRodapés encontrados:"
    doc.footers.each do |key, footer|
      puts "- #{key}: #{footer.text}"
    end
  else
    puts "\nNenhum rodapé encontrado."
  end
end

# Função para substituir o conteúdo de cabeçalhos e rodapés
def replace_headers_footers(doc)
  # Substituir conteúdo dos cabeçalhos
  if doc.headers && !doc.headers.empty?
    doc.headers.each do |key, header|
      # Como parece que não podemos acessar paragraphs diretamente,
      # vamos tentar modificar o XML do cabeçalho diretamente
      begin
        # Encontrar nós de parágrafo no XML
        if header.is_a?(Nokogiri::XML::Document)
          # Encontrar todos os nós de texto
          text_nodes = header.xpath('//w:t')
          
          if text_nodes.any?
            text_nodes.each do |node|
              node.content = "Custom Header"
            end
            puts "Substituído conteúdo do cabeçalho: #{key}"
          else
            # Se não encontrarmos nós de texto, tente criar um novo parágrafo
            paragraphs = header.xpath('//w:p')
            if paragraphs.any?
              # Pegar o primeiro parágrafo e substituir
              p_node = paragraphs.first
              
              # Limpar conteúdo existente
              p_node.children.remove
              
              # Criar e adicionar novo texto
              text_run = Nokogiri::XML::Node.new('w:r', header)
              text_node = Nokogiri::XML::Node.new('w:t', header)
              text_node.content = "Custom Header"
              
              text_run.add_child(text_node)
              p_node.add_child(text_run)
              
              puts "Criado novo texto para cabeçalho: #{key}"
            end
          end
        end
      rescue => e
        puts "Erro ao modificar cabeçalho #{key}: #{e.message}"
      end
    end
  end
  
  # Substituir conteúdo dos rodapés
  if doc.footers && !doc.footers.empty?
    doc.footers.each do |key, footer|
      begin
        if footer.is_a?(Nokogiri::XML::Document)
          # Encontrar todos os nós de texto
          text_nodes = footer.xpath('//w:t')
          
          if text_nodes.any?
            text_nodes.each do |node|
              node.content = "Custom Footer"
            end
            puts "Substituído conteúdo do rodapé: #{key}"
          else
            # Se não encontrarmos nós de texto, tente criar um novo parágrafo
            paragraphs = footer.xpath('//w:p')
            if paragraphs.any?
              # Pegar o primeiro parágrafo e substituir
              p_node = paragraphs.first
              
              # Limpar conteúdo existente
              p_node.children.remove
              
              # Criar e adicionar novo texto
              text_run = Nokogiri::XML::Node.new('w:r', footer)
              text_node = Nokogiri::XML::Node.new('w:t', footer)
              text_node.content = "Custom Footer"
              
              text_run.add_child(text_node)
              p_node.add_child(text_run)
              
              puts "Criado novo texto para rodapé: #{key}"
            end
          end
        end
      rescue => e
        puts "Erro ao modificar rodapé #{key}: #{e.message}"
      end
    end
  end
  
  # Salvar os documentos modificados de volta no docx
  # Isso só funcionará se o PR implementou um método para acessar
  # o conteúdo XML do header/footer e salvá-lo de volta
  if doc.respond_to?(:save_headers_and_footers)
    doc.save_headers_and_footers
    puts "Headers e footers salvos de volta no documento"
  end
end

begin
  # Verificar se o arquivo existe
  unless File.exist?(DOCX_PATH)
    puts "Erro: Arquivo #{DOCX_PATH} não encontrado."
    exit 1
  end
  
  # Abrir o documento
  doc = Docx::Document.open(DOCX_PATH)
  
  # Mostrar informações antes da modificação
  puts "=== ANTES DA MODIFICAÇÃO ==="
  show_doc_info(doc)
  
  # Modificar cabeçalhos e rodapés
  replace_headers_footers(doc)
  
  # Salvar o documento modificado
  new_file_path = "modified_#{File.basename(DOCX_PATH)}"
  doc.save(new_file_path)
  puts "\nDocumento salvo como: #{new_file_path}"
  
  # Abrir o novo documento para verificar mudanças
  new_doc = Docx::Document.open(new_file_path)
  
  # Mostrar informações após a modificação
  puts "\n=== APÓS A MODIFICAÇÃO ==="
  show_doc_info(new_doc)
  
rescue StandardError => e
  puts "Erro: #{e.message}"
  puts e.backtrace
end