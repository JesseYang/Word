# encoding: utf-8
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
java_import 'com.aspose.words.SmartTag'
java_import 'com.aspose.words.Run'
java_import 'com.aspose.words.OfficeMath'
java_import 'com.aspose.words.Shape'
java_import 'com.aspose.words.Table'
java_import 'com.aspose.words.DrawingML'
java_import 'com.aspose.words.WrapType'
java_import 'com.aspose.words.VerticalAlignment'
java_import 'org.apache.xmlbeans.XmlObject'
java_import 'org.apache.poi.xwpf.usermodel.XWPFDocument'
java_import 'java.io.InputStream'
java_import 'java.io.FileInputStream'
java_import 'java.io.OutputStream'
java_import 'java.io.FileOutputStream'

configure do
  set :item_thresh, [10, 20]
  set :line_length, 80
  set :suffix_ary, ["", "", "emf", "wmf", "pict", "jpeg", "png", "bmp"]
end

# http://www.aspose.com/docs/display/wordsjava/ImageType

get '/extract' do
  filename = "#{settings.root}/../EngLib/public/uploads/documents/#{params[:filename]}"
  filename_ary = split_doc(filename)
  content = []
  filename_ary.each do |f|
    doc = Document.new(f)
    doc.sections.get(0).body.getChildNodes.each_with_index do |node, index|
      next if f.include?("true") && index == doc.sections.get(0).body.getChildNodes.count - 1
      if node.class == Paragraph
        string = parse_paragraph(node).select { |e| !e.start_with?("Evaluation Only.") }
        content += string
      elsif node.class == Table
        content << parse_table(node)
      elsif node.class == Shape
      end
    end
    File.delete(f)
  end
  content_type :json
    { content: content }.to_json
end

post '/export_note' do
  params = JSON.parse(request.body.read)
  doc = Document.new("template.docx")
  builder = DocumentBuilder.new(doc)
  builder.moveToDocumentEnd()
  params["questions"].each do |q|
    q["content"].each do |node|
      if node.class == String
        # normal line
        write_paragraph(builder, node)
      elsif node.class == Hash && node["type"] == "table"
        # table
        write_table(builder, node)
      end
    end
    if q["type"] == "choice"
      organize_items(q["items"]).each { |e| write_paragraph(builder, e) }
    end
    
    if (q["answer_content"] || []).length > 0 || !q["answer"].nil?
        write_answer(builder, q)
        2.times { builder.writeln("") }
    else
      if q["type"] == "choice"
        2.times { builder.writeln("") }
      else
        10.times {builder.writeln("") }
      end
    end
  end
  filename = "downloads/documents/#{params["name"]}_#{SecureRandom.uuid}.docx"
  fullpath = "#{settings.root}/../EngLib/public/#{filename}"
  doc.save(fullpath)
  content_type :json
    { filename: remove_ad(fullpath, params["name"]) }.to_json
end

post '/generate' do
  params = JSON.parse(request.body.read)
  doc = Document.new("template.docx")
  builder = DocumentBuilder.new(doc)
  builder.moveToDocumentEnd()
  params["questions"].each do |q|

    qr = RQRCode::QRCode.new(q["link"], :size => 4, :level => :l )
    png = qr.to_img
    temp_img_name = "public/#{SecureRandom.uuid}.png"
    png.resize(70, 70).save(temp_img_name)
    shape = builder.insertImage(temp_img_name)
    shape.setWrapType(WrapType::SQUARE)
    shape.setLeft(370)

    q["content"].each do |node|
      if node.class == String
        # normal line
        write_paragraph(builder, node)
      elsif node.class == Hash && node["type"] == "table"
        # table
        write_table(builder, node)
      end
    end
    if q["type"] == "choice"
      organize_items(q["items"]).each { |e| write_paragraph(builder, e) }
    end
    
    if (q["answer_content"] || []).length > 0 || !q["answer"].nil?
        write_answer(builder, q)
        2.times { builder.writeln("") }
    else
      if q["type"] == "choice"
        2.times { builder.writeln("") }
      else
        10.times {builder.writeln("") }
      end
    end
  end
  filename = "downloads/documents/#{params["name"]}_#{SecureRandom.uuid}.docx"
  fullpath = "#{settings.root}/../EngLib/public/#{filename}"
  doc.save(fullpath)
  content_type :json
    { filename: remove_ad(fullpath, params["name"]) }.to_json
end

post '/export' do
  params = JSON.parse(request.body.read)
  doc = Document.new("template.docx")
  builder = DocumentBuilder.new(doc)
  builder.moveToDocumentEnd()
  params["groups"].each do |questions|
    questions.each do |q|
      # write content
      q["content"].each do |node|
        if node.class == String
          # normal line
          write_paragraph(builder, node)
        elsif node.class == Hash && node["type"] == "table"
          # table
          write_table(builder, node)
        end
      end
      # write items
      if q["type"] == "choice"
        organize_items(q["items"]).each { |e| write_paragraph(builder, e) }
      end
      # write answers
      if (q["answer_content"] || []).length > 0 || !q["answer"].nil?
        write_answer(builder, q)
      end
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

def parse_table(table)
  content = []
  parsed_table = {type: "table", content: content}
  table.getChildNodes.each do |row|
    parsed_row = []
    row.getChildNodes.each do |cell|
      parsed_cell = []
      cell.each do |para|
        parsed_cell += parse_paragraph(para)
      end
      parsed_row << parsed_cell
    end
    content << parsed_row
  end
  parsed_table
end

def parse_paragraph(para)
  para_text = []
  cur_text = ""
  para.getChildNodes.each do |e|
    type_info = judge_type(e)
    case type_info[0]
    when "text"
      cur_text += e.text
    when "unknown"
      cur_text += e.text
    when "equation"
      suffix = settings.suffix_ary[e.getImageData().imageType]
      next if suffix == ""
      img_file_name = "#{SecureRandom.uuid}.#{suffix}"
      cur_text += "$equation*#{img_file_name}*#{type_info[1].to_s}*#{type_info[2].to_s}$"
      e.getImageData().save("#{settings.root}/../EngLib/public/uploads/documents/images/#{img_file_name}")
    when "figure"
      para_text << cur_text if cur_text != ""
      cur_text = ""
      suffix = settings.suffix_ary[e.getImageData().imageType]
      next if suffix == ""
      img_file_name = "#{SecureRandom.uuid}.#{suffix}"
      para_text << "$figure*#{img_file_name}*#{type_info[1].to_s}*#{type_info[2].to_s}$"
      e.getImageData().save("#{settings.root}/../EngLib/public/uploads/documents/images/#{img_file_name}")
    end
  end
  para_text << cur_text if cur_text != ""
  para_text << "" if para_text == []
  para_text
end

def judge_type(e)
  if e.class == Run || e.class == SmartTag
    ["text"]
  elsif e.class == Shape && settings.suffix_ary[e.getImageData().imageType] == "wmf"
    ["equation", e.getWidth, e.getHeight]
  elsif (e.class == DrawingML || e.class == Shape) && e.isInline == false
    ["figure", e.getWidth, e.getHeight]
  elsif (e.class == DrawingML || e.class == Shape) && e.isInline
    ["equation", e.getWidth, e.getHeight]
  else
    ["unknown"]
  end
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

def write_answer(builder, question)
  builder.writeln("答案:#{d2c(question["answer"])}")
  question["answer_content"].each do |node|
    if node.class == String
      # normal line
      write_paragraph(builder, node)
    elsif node.class == Hash && node["type"] == "table"
      # table
      write_table(builder, node)
    end
  end
end

# http://www.aspose.com/docs/display/wordsjava/Inserting+a+Table+using+DocumentBuilder
def write_table(builder, table)
  builder.startTable()
  table["content"].each do |row|
    row.each do |cell|
      builder.insertCell()
      cell.each do |para|
        write_paragraph(builder, para, false)
      end
    end
    builder.endRow()
  end
  builder.endTable()
end

def write_paragraph(builder, content, new_line = true)
  content.split('$').each do |f|
    if f.match(/[a-z 0-9]{8}-[a-z 0-9]{4}-[a-z 0-9]{4}-[a-z 0-9]{4}-[a-z 0-9]{12}/)
      # equation
      shape = builder.insertImage(("#{settings.root}/../EngLib/public/uploads/documents/images/#{f}"))
      shape.setWrapType(WrapType::INLINE)
      shape.setVerticalAlignment(VerticalAlignment::CENTER)
    else
      # text
      builder.write(f)
    end
  end
  builder.writeln("") if new_line
end

def d2c(d)
  return "" if d.nil?
  ("A".."Z").to_a[d]
end

def split_doc(filename)
  prefix = SecureRandom.uuid
  filename_ary = []
  doc = XWPFDocument.new(FileInputStream.new(filename))
  elementNumber = doc.getBodyElements.length
  all_index = (0..elementNumber-1).to_a
  # every 50 elements are split as one doc
  size = 50
  file_num = (elementNumber * 1.0 / size).ceil
  (0..file_num - 1).to_a.each do |file_index|
    remove_pre_index_ary = all_index.select do |e|
      e < file_index * size
    end
    remove_aft_index_ary = all_index.select do |e|
      e >= (file_index + 1) * size
    end
    doc = XWPFDocument.new(FileInputStream.new(filename))
    remove_aft_index_ary.reverse.each do |e|
      doc.removeBodyElement(e)
    end
    remove_pre_index_ary.each do |e|
      doc.removeBodyElement(0)
    end
    delete = doc.getBodyElements[doc.getBodyElements.length-1].class == Java::OrgApachePoiXwpfUsermodel::XWPFTable
    split_f = "#{prefix}-#{file_index}-#{delete}.docx"
    doc.write(FileOutputStream.new(split_f))
    filename_ary << split_f
  end
  filename_ary
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
