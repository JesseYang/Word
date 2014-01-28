require 'rubygems'
require 'securerandom'
require 'json'
require 'sinatra'
require 'rqrcode_png'
require 'java'
require 'lib/Aspose.Words.jdk16'
require 'lib/poi-3.9-20121203'
require 'lib/poi-ooxml-3.9-20121203'
require 'lib/xmlbeans-2.3.0'
require 'lib/poi-ooxml-schemas-3.9-20121203'
require 'lib/poi-scratchpad-3.9-20121203'
require 'lib/dom4j-1.6.1'
require 'lib/stax-api-1.0.1'
java_import 'com.aspose.words.Document'
java_import 'com.aspose.words.DocumentBuilder'
java_import 'com.aspose.words.Section'
java_import 'com.aspose.words.Body'
java_import 'com.aspose.words.Paragraph'
java_import 'com.aspose.words.Run'
java_import 'org.apache.xmlbeans.XmlObject'
java_import 'org.apache.poi.xwpf.usermodel.XWPFDocument'
java_import 'java.io.InputStream'
java_import 'java.io.FileInputStream'
java_import 'java.io.OutputStream'
java_import 'java.io.FileOutputStream'

configure do
  set :item_thresh, [10, 20]
  set :line_length, 80
end

get '/extract' do
  filename = "#{settings.root}/../EngLib/public/uploads/documents/#{params[:filename]}"
  doc = Document.new(filename)
  content = []
  doc.sections.get(0).body.paragraphs.each do |para|
    content << (para.runs.map { |e| e.text } .join)
  end
  content_type :json
    { content: content }.to_json
end

post '/generate' do
  params = JSON.parse(request.body.read)
  doc = Document.new
  builder = DocumentBuilder.new(doc)
  params["questions"].each do |q|
    builder.writeln(q["content"])
    organize_items(q["items"]).each { |e| builder.writeln(e) }
    qr = RQRCode::QRCode.new(q["link"], :size => 4, :level => :l )
    png = qr.to_img
    temp_img_name = "public/#{SecureRandom.uuid}.png"
    png.resize(90, 90).save(temp_img_name)
    builder.insertImage(temp_img_name)
    builder.writeln("")
  end
  filename = "downloads/documents/#{params["name"]}_#{SecureRandom.uuid}.docx"
  fullpath = "#{settings.root}/../EngLib/public/#{filename}"
  doc.save(fullpath)
  content_type :json
    { filename: remove_ad(fullpath, params["name"]) }.to_json
end

post '/export' do
  params = JSON.parse(request.body.read)
  doc = Document.new
  builder = DocumentBuilder.new(doc)
  params["groups"].each do |questions|
    questions.each do |q|
      builder.writeln(q["content"])
      organize_items(q["items"]).each { |e| builder.writeln(e) }
      builder.writeln("")
    end
    builder.writeln("-" * 60)
    builder.writeln("")
  end
  filename = "downloads/documents/#{params["name"]}_#{SecureRandom.uuid}.docx"
  fullpath = "#{settings.root}/../EngLib/public/#{filename}"
  doc.save(fullpath)
  content_type :json
    { filename: remove_ad(fullpath, params["name"]) }.to_json
end

def organize_items(items)
  # plus prefix for each item
  items = items.each_with_index.map do |e, i|
    "#{("A".."Z").to_a[i]}. #{e}"
  end
  max_length = items.map { |e| e.length } .max
  if max_length < settings.item_thresh[0]
    # should in one line
    occ_len = settings.line_length/4
    line = "#{items[0]}#{" " * (occ_len-items[0].length)}"
    line += "#{items[1]}#{" " * (occ_len-items[1].length)}"
    line += "#{items[2]}#{" " * (occ_len-items[2].length)}"
    line += "#{items[3]}"
    return [line]
  elsif max_length < settings.item_thresh[1]
    # should in two lines
    lines = []
    occ_len = settings.line_length/2
    lines << "#{items[0]}#{" " * (occ_len-items[0].length)}#{items[1]}"
    lines << "#{items[2]}#{" " * (occ_len-items[2].length)}#{items[3]}"
    return lines
  else
    # each item occupies one line
   return items
  end
end

def remove_ad(temp_file_name, original_name)
  doc = XWPFDocument.new(FileInputStream.new(temp_file_name))
  doc.paragraphs[0].runs[0].setText("", 0)
  filename = "downloads/documents/#{original_name}_#{SecureRandom.uuid}.docx"
  fullpath = "#{settings.root}/../EngLib/public/#{filename}"
  file_out = FileOutputStream.new(fullpath)
  doc.write(file_out)
  filename 
end
