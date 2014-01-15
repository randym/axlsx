unless String.method_defined? :prepend
  class String
    def prepend(other_str)
      insert(0, other_str)
    end
  end
end