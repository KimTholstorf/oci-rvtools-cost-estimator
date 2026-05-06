class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/55/0c/b847c3591b562221c657ba39de62450b7880a7406805ba5d94c98f726f7a/oci_rvtools-1.0.11.tar.gz"
  sha256 "12c666b68503fbb197e05bea41af2e19f3008ee1ee0c980d1f289895cc8b933a"
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
