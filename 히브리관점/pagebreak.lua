function HorizontalRule()
  if FORMAT:match('docx') then
    return pandoc.RawBlock('openxml', '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
  end

  return pandoc.RawBlock('html', '<div style="page-break-after: always;"></div>')
end
