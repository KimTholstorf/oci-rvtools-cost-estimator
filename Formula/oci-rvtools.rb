class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/cf/68/92a483a7a99d2d5913b04f0642c891930771b8430cba78c80ed3853808ec/oci_rvtools-1.0.9.tar.gz"
  sha256 "b214136a38b9e67657c57f64fcef7191b1fb91287173e2193ab0b60b71b5e83d"
  license "MIT"

  depends_on "python3"

  def install
    system "python3", "-m", "venv", libexec
    system libexec/"bin/pip", "install", "--no-cache-dir", "oci-rvtools==#{version}"
    bin.install_symlink libexec/"bin/oci-rvtools"
  end

  test do
    system bin/"oci-rvtools", "--version"
  end
end
