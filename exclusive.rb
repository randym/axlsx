
def exclusively_true(hash, key)
  #clone = hash.clone
  return false unless hash.delete(key) == true
  !hash.has_value? true
end

require 'test/unit'
class TestExclusive < Test::Unit::TestCase
  def setup
    @test_hash = {foo: true, bar: false, hoge: false}
  end
  def test_exclusive
    assert_equal(true, exclusively_true(@test_hash, :foo))
  end
  def test_inexclusive
    @test_hash[:bar] = true
    assert_equal(false, exclusively_true(@test_hash, :foo))
  end
end

require 'benchmark'

h = {foo: true}
999.times {|i| h["a#{i}"] = false}
Benchmark.bmbm(30) do |x|
  x.report('exclusively_true') do
    1000.times do
      exclusively_true(h, :foo)
    end
  end
end
