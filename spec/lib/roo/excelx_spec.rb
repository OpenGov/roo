require 'spec_helper'

describe Roo::Excelx do
  describe '.new' do
    subject {
      Roo::Excelx.new('test/files/numbers1.xlsx')
    }

    it 'creates an instance' do
      expect(subject).to be_a(Roo::Excelx)
    end

    it 'correctly parses files that dont have a spans attribute on rows' do
      parsed = Roo::Excelx.new('test/files/missing_spans.xlsx')
      expect(parsed.row(1)).to eql(["Adulterated document"])
    end
  end
end
