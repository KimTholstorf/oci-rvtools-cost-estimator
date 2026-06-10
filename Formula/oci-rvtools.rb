class OciRvtools < Formula
  desc "Convert RVTools Excel exports into an Oracle Cloud (OCI) monthly cost estimate workbook"
  homepage "https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
  url "https://files.pythonhosted.org/packages/f5/8b/ee7b03c5e5216f0c8d459c07040bf6a6dd5c3d6b7b2cbc20e13425a536cb/oci_rvtools-1.2.1.tar.gz"
  sha256 "d1f0d3422a5a65cf6189a3df6fa149055fb0a17c3770ab40f115305717bab7c2"
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
