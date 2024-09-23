-- Based on https://gist.github.com/const-ae/752ad85c43d92b72865453ea3a77e2dd#file-resolve_equation_labels-lua
-- Bail if we are converting to LaTeX.
if FORMAT:match('latex') then
  return {}
end

local nequations = 0
local equation_labels = {}

-- Helper function to find and remove \label from the equation text
local function find_label(txt)
  local before, label, after = txt:match('(.*)\\label%{(.-)%}(.*)')
  return label, label and before .. after
end

return {
  {
    -- Process math blocks to track equation labels
    Math = function(m)
      if m.mathtype == pandoc.DisplayMath then
        nequations = nequations + 1
        local label, stripped = find_label(m.text)
        if label then
          equation_labels[label] = "(" .. tostring(nequations) .. ")"
        end
      end
      return m
    end
  },
  {
    -- Handle equation references
    Link = function(link)
      local ref = link.attributes.reference
      if ref and link.attributes['reference-type'] == 'eqref' and equation_labels[ref] then
        link.content = pandoc.Str(equation_labels[ref])
      end
      return link
    end
  }
}
