require 'bundler/setup'
require 'docx'
require 'zip'
require 'fileutils'
require 'mini_magick'

# Create temp directory for extraction
temp_dir = "temp_docx_extract"
FileUtils.mkdir_p(temp_dir)
FileUtils.rm_rf(temp_dir) # Clean up any previous extraction
FileUtils.mkdir_p(temp_dir)

# Extract the docx
input_file = "procuracao.docx"
Zip::File.open(input_file) do |zip_file|
  zip_file.each do |entry|
    entry_path = File.join(temp_dir, entry.name)
    FileUtils.mkdir_p(File.dirname(entry_path))
    zip_file.extract(entry, entry_path) unless File.exist?(entry_path)
  end
end

# Process image1.png (footer)
image_path = File.join(temp_dir, "word/media/image1.png")
if File.exist?(image_path)
  image = MiniMagick::Image.open(image_path)

  # Get image metadata and attributes
  width = image.width
  height = image.height
  size = File.size(image_path)
  metadata = image.exif rescue {}
  puts "Original image1 attributes: #{width}x#{height}, #{size} bytes"
  puts "Metadata: #{metadata}"

  # Process the new image to match the original
  new_image = MiniMagick::Image.open("image5.png")
  new_image.resize "#{width}x#{height}"
  
  # Copy metadata if needed
  metadata.each do |key, value|
    new_image.exif[key] = value if new_image.exif
  end

  # Save the modified new image to the temp directory
  new_image.write(File.join(temp_dir, "word/media/image1.png"))
else
  puts "Warning: image1.png not found at expected path: #{image_path}"
end

# Process image2.png (header)
image_path = File.join(temp_dir, "word/media/image2.png")
if File.exist?(image_path)
  image = MiniMagick::Image.open(image_path)

  # Get image metadata and attributes
  width = image.width
  height = image.height
  size = File.size(image_path)
  metadata = image.exif rescue {}
  puts "Original image2 attributes: #{width}x#{height}, #{size} bytes"
  puts "Metadata: #{metadata}"

  # Process the new image to match the original
  new_image = MiniMagick::Image.open("image6.png")
  new_image.resize "#{width}x#{height}"
  
  # Copy metadata if needed
  metadata.each do |key, value|
    new_image.exif[key] = value if new_image.exif
  end

  # Save the modified new image to the temp directory
  new_image.write(File.join(temp_dir, "word/media/image2.png"))
else
  puts "Warning: image2.png not found at expected path: #{image_path}"
end

# Create the new docx file by directly copying files without using the normal zip methods
output_file = "procuracao_modified.docx"
FileUtils.rm(output_file) if File.exist?(output_file)

# Use the -r option for recursion
system("cd #{temp_dir} && zip -r ../#{output_file} *")

# Clean up
FileUtils.rm_rf(temp_dir)

puts "DOCX file successfully modified with new header and footer images."