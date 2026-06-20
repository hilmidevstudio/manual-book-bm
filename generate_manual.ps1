# Script untuk melanjutkan Manual Book BuildMatch
# Menambahkan bagian Client, Kontraktor, dan Arsitek

$sourceDocx = "MANUAL BOOK BUILDMATCH.docx"
$outputDocx = "MANUAL BOOK BUILDMATCH - LENGKAP.docx"
$tempDir = "docx-build-temp"

Write-Host "=== Memulai proses generate manual book ===" -ForegroundColor Cyan

# Buat direktori temp
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempDir | Out-Null

# Copy file asli dan extract
Copy-Item $sourceDocx "$tempDir\source.zip"
Expand-Archive "$tempDir\source.zip" -DestinationPath "$tempDir\docx" -Force

Write-Host "File asli berhasil diekstrak" -ForegroundColor Green

# Daftar gambar yang akan ditambahkan
$imagesToAdd = @()

# CLIENT images (folder 2-client)
$clientImages = @(
    "2-client\1-sudah-login-client.jpeg",
    "2-client\2-menu-utama-beranda.jpeg",
    "2-client\3-mitra1-kontraktor.jpeg",
    "2-client\3-mitra2-arsitek.jpeg",
    "2-client\4-pre-konsultasi1-kontraktor.jpeg",
    "2-client\4-pre-konsultasi2-arsitek.jpeg",
    "2-client\4-roomchat-konsultasi1-kontraktor.jpeg",
    "2-client\5-progress1-pembangunan1.jpeg",
    "2-client\5-progress1-pembangunan2-detail-proyek1.jpeg",
    "2-client\5-progress1-pembangunan2-detail-proyek2.jpeg",
    "2-client\5-progress-pembangunan2-detail-proyek3-detail-penawaran.jpeg",
    "2-client\5-progress1-pembangunan2-detail-proyek4-termin-pembayaran1.jpeg",
    "2-client\5-progress1-pembangunan2-detail-proyek4-termin-pembayaran2.jpeg",
    "2-client\5-progress2-desain1.jpeg",
    "2-client\5-progress2-desain1-detail-desain.jpeg",
    "2-client\5-progress3-draft.jpeg",
    "2-client\6-progress-update-draft.jpeg",
    "2-client\6-profile1.jpeg",
    "2-client\6-profile1-edit-profile.jpeg",
    "2-client\6-profile2.jpeg",
    "2-client\6-profile2-buatproyek1.jpeg",
    "2-client\6-profile2-buatproyek2-1.jpeg",
    "2-client\6-profile2-buatproyek2-2.jpeg",
    "2-client\6-profile2-buatproyek3-1.jpeg",
    "2-client\6-profile2-buatproyek3-2.jpeg",
    "2-client\6-profile2-buatproyek4-1.jpeg",
    "2-client\6-profile2-buatproyek4-2.jpeg",
    "2-client\6-profile2-buatproyek4-3.jpeg",
    "2-client\6-profile2-buatproyek5-notif.jpeg",
    "2-client\6-profile3-update-profile.jpeg"
)

# KONTRAKTOR images (folder 3-kontraktor)
$kontraktorImages = @(
    "3-kontraktor\1-sudah-login-client.jpeg",
    "3-kontraktor\2-penawaran-masuk1-detail-proyek1.jpeg",
    "3-kontraktor\2-penawaran-masuk2-diajukan1.jpeg",
    "3-kontraktor\2-penawaran-masuk2-diajukan2.jpeg",
    "3-kontraktor\2-penawaran-masuk2-diajukan3.jpeg",
    "3-kontraktor\3-progress1.jpeg",
    "3-kontraktor\3-progress1-detail-penawaran1.jpeg",
    "3-kontraktor\3-progress1-detail-penawaran2.jpeg",
    "3-kontraktor\3-progress1-detail-penawaran3-termin-pembayaran1.jpeg",
    "3-kontraktor\3-progress1-detail-penawaran3-termin-pembayaran2.jpeg",
    "3-kontraktor\4-edit-profil1.jpeg",
    "3-kontraktor\4-edit-profil2.jpeg",
    "3-kontraktor\4-edit-profil3-1.jpeg",
    "3-kontraktor\4-edit-profil3-2.jpeg",
    "3-kontraktor\4-edit-profil3-3.jpeg",
    "3-kontraktor\4-edit-profil4-notif.jpeg"
)

# ARSITEK images (folder 4-arsitek)
$arsitekImages = @(
    "4-arsitek\1-sudah-login-arsitek1.jpeg",
    "4-arsitek\2-desain1-upload-desain1.jpeg",
    "4-arsitek\2-desain1-upload-desain2.jpeg",
    "4-arsitek\3-inbox1-belum-terbaca.jpeg",
    "4-arsitek\3-inbox2-roomchat1-1.jpeg",
    "4-arsitek\3-inbox2-roomchat1-2.jpeg",
    "4-arsitek\3-inbox3-sudah-terbaca.jpeg",
    "4-arsitek\4-profil1.jpeg",
    "4-arsitek\4-profil2.jpeg",
    "4-arsitek\4-profil3.jpeg",
    "4-arsitek\4-profil5-edit-profil1.jpeg",
    "4-arsitek\4-profil5-edit-profl2.jpeg",
    "4-arsitek\4-profil5-edit-profil3.jpeg"
)

$allImages = $clientImages + $kontraktorImages + $arsitekImages

# Copy semua gambar baru ke folder media
$mediaDir = "$tempDir\docx\word\media"
$startRId = 30  # rId30 dan seterusnya

$imageRIds = @{}
$imageIndex = 21  # image21.jpeg dan seterusnya (setelah image20.jpeg yang sudah ada)

foreach ($imgPath in $allImages) {
    if (Test-Path $imgPath) {
        $ext = [System.IO.Path]::GetExtension($imgPath)
        $newName = "image$imageIndex$ext"
        Copy-Item $imgPath "$mediaDir\$newName"
        $rId = "rId$startRId"
        $imageRIds[$imgPath] = @{ RId = $rId; FileName = $newName; Index = $imageIndex }
        $startRId++
        $imageIndex++
        Write-Host "  + $imgPath -> $rId ($newName)" -ForegroundColor Gray
    } else {
        Write-Warning "File tidak ditemukan: $imgPath"
    }
}

Write-Host "Semua gambar berhasil disalin ($($imageRIds.Count) gambar)" -ForegroundColor Green

# Update relationships file
$relsFile = "$tempDir\docx\word\_rels\document.xml.rels"
$relsContent = Get-Content $relsFile -Raw

# Tambahkan relationships baru sebelum </Relationships>
$newRels = ""
foreach ($imgPath in $allImages) {
    if ($imageRIds.ContainsKey($imgPath)) {
        $info = $imageRIds[$imgPath]
        $newRels += "<Relationship Id=`"$($info.RId)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/$($info.FileName)`"/>"
    }
}
$relsContent = $relsContent -replace "</Relationships>", "$newRels</Relationships>"
Set-Content $relsFile $relsContent -Encoding UTF8

Write-Host "Relationships file berhasil diupdate" -ForegroundColor Green

# Helper function untuk membuat XML paragraf gambar (portrait/mobile screenshot)
function New-ImageParagraph {
    param($rId, $imgId, $imgName, [int]$widthEmu = 1346200, [int]$heightEmu = 2979040)
    return @"
<w:p w14:paraId="IMG$imgId"><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr><w:bidi w:val="0"/><w:ind w:left="1320" w:leftChars="0" w:right="0" w:rightChars="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:hint="default"/><w:lang w:val="id-ID"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="default"/><w:lang w:val="id-ID"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="114300" distR="114300"><wp:extent cx="$widthEmu" cy="$heightEmu"/><wp:effectExtent l="0" t="0" r="9525" b="4445"/><wp:docPr id="$imgId" name="Picture $imgId" descr="$imgName"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$imgId" name="Picture $imgId" descr="$imgName"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="$rId"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="$widthEmu" cy="$heightEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
"@
}

# Helper function untuk dua gambar side-by-side
function New-TwoImageParagraph {
    param($rId1, $imgId1, $imgName1, $rId2, $imgId2, $imgName2, [int]$widthEmu = 1346200, [int]$heightEmu = 2979040)
    return @"
<w:p w14:paraId="TWO$imgId1"><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr><w:bidi w:val="0"/><w:ind w:left="1320" w:leftChars="0" w:right="0" w:rightChars="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:hint="default"/><w:lang w:val="id-ID"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="default"/><w:lang w:val="id-ID"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="114300" distR="228600"><wp:extent cx="$widthEmu" cy="$heightEmu"/><wp:effectExtent l="0" t="0" r="9525" b="4445"/><wp:docPr id="$imgId1" name="Picture $imgId1" descr="$imgName1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$imgId1" name="Picture $imgId1" descr="$imgName1"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="$rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="$widthEmu" cy="$heightEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing><w:drawing><wp:inline distT="0" distB="0" distL="228600" distR="114300"><wp:extent cx="$widthEmu" cy="$heightEmu"/><wp:effectExtent l="0" t="0" r="9525" b="4445"/><wp:docPr id="$imgId2" name="Picture $imgId2" descr="$imgName2"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$imgId2" name="Picture $imgId2" descr="$imgName2"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="$rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="$widthEmu" cy="$heightEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
"@
}

# Helper functions untuk paragraf teks
function New-HeadingL1 { param($text, $bookmarkId, $bookmarkName)
    return "<w:p w14:paraId=`"H1$(Get-Random -Max 9999999)`"><w:pPr><w:pStyle w:val=`"2`"/><w:numPr><w:ilvl w:val=`"0`"/><w:numId w:val=`"2`"/></w:numPr><w:bidi w:val=`"0`"/><w:ind w:left=`"0`" w:leftChars=`"0`" w:firstLine=`"0`" w:firstLineChars=`"0`"/><w:rPr><w:rFonts w:hint=`"default`"/><w:lang w:val=`"id-ID`"/></w:rPr></w:pPr><w:bookmarkStart w:id=`"$bookmarkId`" w:name=`"$bookmarkName`"/><w:r><w:rPr><w:rFonts w:hint=`"default`"/><w:lang w:val=`"id-ID`"/></w:rPr><w:t>$text</w:t></w:r><w:bookmarkEnd w:id=`"$bookmarkId`"/></w:p>"
}

function New-HeadingL2 { param($text, $bookmarkId, $bookmarkName)
    return "<w:p w14:paraId=`"H2$(Get-Random -Max 9999999)`"><w:pPr><w:pStyle w:val=`"2`"/><w:numPr><w:ilvl w:val=`"1`"/><w:numId w:val=`"2`"/></w:numPr><w:bidi w:val=`"0`"/><w:rPr><w:rFonts w:hint=`"default`"/><w:lang w:val=`"id-ID`"/></w:rPr></w:pPr><w:bookmarkStart w:id=`"$bookmarkId`" w:name=`"$bookmarkName`"/><w:r><w:rPr><w:rFonts w:hint=`"default`"/><w:lang w:val=`"id-ID`"/></w:rPr><w:t>$text</w:t></w:r><w:bookmarkEnd w:id=`"$bookmarkId`"/></w:p>"
}

function New-HeadingL3 { param($text)
    return "<w:p w14:paraId=`"H3$(Get-Random -Max 9999999)`"><w:pPr><w:numPr><w:ilvl w:val=`"2`"/><w:numId w:val=`"2`"/></w:numPr><w:bidi w:val=`"0`"/><w:rPr><w:rFonts w:hint=`"default`"/><w:lang w:val=`"id-ID`"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=`"default`"/><w:lang w:val=`"id-ID`"/></w:rPr><w:t>$text</w:t></w:r></w:p>"
}

function New-BoldPara { param($text)
    return "<w:p w14:paraId=`"BP$(Get-Random -Max 9999999)`"><w:pPr><w:pStyle w:val=`"12`"/><w:keepNext w:val=`"0`"/><w:keepLines w:val=`"0`"/><w:widowControl/><w:suppressLineNumbers w:val=`"0`"/><w:ind w:left=`"1320`" w:leftChars=`"600`" w:firstLine=`"0`" w:firstLineChars=`"0`"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>$text</w:t></w:r></w:p>"
}

function New-TextPara { param($text)
    return "<w:p w14:paraId=`"TP$(Get-Random -Max 9999999)`"><w:pPr><w:pStyle w:val=`"12`"/><w:keepNext w:val=`"0`"/><w:keepLines w:val=`"0`"/><w:widowControl/><w:suppressLineNumbers w:val=`"0`"/><w:ind w:left=`"1320`" w:leftChars=`"600`" w:firstLine=`"0`" w:firstLineChars=`"0`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$text</w:t></w:r></w:p>"
}

function New-NumberedStep { param($text)
    return "<w:p w14:paraId=`"NS$(Get-Random -Max 9999999)`"><w:pPr><w:pStyle w:val=`"12`"/><w:keepNext w:val=`"0`"/><w:keepLines w:val=`"0`"/><w:widowControl/><w:numPr><w:ilvl w:val=`"0`"/><w:numId w:val=`"3`"/></w:numPr><w:suppressLineNumbers w:val=`"0`"/><w:ind w:left=`"1587`" w:leftChars=`"0`" w:hanging=`"267`" w:firstLineChars=`"0`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$text</w:t></w:r></w:p>"
}

function New-BodyIndent { param($text)
    return "<w:p w14:paraId=`"BI$(Get-Random -Max 9999999)`"><w:pPr><w:bidi w:val=`"0`"/><w:ind w:left=`"720`" w:leftChars=`"0`" w:firstLine=`"720`" w:firstLineChars=`"0`"/><w:rPr><w:sz w:val=`"24`"/><w:szCs w:val=`"24`"/></w:rPr></w:pPr><w:r><w:rPr><w:sz w:val=`"24`"/><w:szCs w:val=`"24`"/></w:rPr><w:t xml:space=`"preserve`">$text</w:t></w:r></w:p>"
}

function New-PageBreak {
    return "<w:p w14:paraId=`"PB$(Get-Random -Max 9999999)`"><w:pPr><w:pStyle w:val=`"9`"/><w:spacing w:before=`"4`"/><w:jc w:val=`"center`"/><w:rPr><w:rFonts w:hint=`"default`"/><w:b/><w:bCs/><w:sz w:val=`"24`"/><w:szCs w:val=`"24`"/><w:lang w:val=`"id-ID`"/></w:rPr><w:sectPr><w:pgSz w:w=`"11920`" w:h=`"16850`"/><w:pgMar w:top=`"0`" w:right=`"1701`" w:bottom=`"0`" w:left=`"1701`" w:header=`"0`" w:footer=`"0`" w:gutter=`"0`"/><w:cols w:space=`"0`" w:num=`"1`"/><w:titlePg/><w:rtlGutter w:val=`"0`"/><w:docGrid w:linePitch=`"0`" w:charSpace=`"0`"/></w:sectPr></w:pPr></w:p>"
}

function New-EmptyPara {
    return "<w:p w14:paraId=`"EP$(Get-Random -Max 9999999)`"><w:pPr><w:bidi w:val=`"0`"/><w:ind w:left=`"720`" w:leftChars=`"0`" w:firstLine=`"720`" w:firstLineChars=`"0`"/></w:pPr></w:p>"
}

Write-Host "Membuat konten XML baru..." -ForegroundColor Cyan

# Lookup helper
function Get-RId { param($imgPath)
    if ($imageRIds.ContainsKey($imgPath)) { return $imageRIds[$imgPath].RId }
    return $null
}

# ID counter
$idCounter = 200

function Next-Id { 
    $script:idCounter++
    return $script:idCounter 
}

# ============================================================
# KONTEN BARU: CLIENT
# ============================================================
$clientContent = ""

# Section break / heading
$clientContent += New-PageBreak
$clientContent += New-HeadingL2 -text "Panduan Pengguna Client (Pemilik Proyek)" -bookmarkId (Next-Id) -bookmarkName "_Toc_Client01"

# 1. Beranda Client
$clientContent += New-HeadingL3 -text "Halaman Beranda Client"
$clientContent += New-TextPara -text "Setelah berhasil masuk ke akun, Client akan diarahkan ke halaman Beranda yang menjadi pusat kendali utama. Halaman ini menampilkan ringkasan aktivitas proyek, menu navigasi utama, daftar mitra terpopuler, dan proyek yang sedang berjalan."
$clientContent += New-BoldPara -text "Komponen utama halaman Beranda:"

$id1 = Next-Id
$clientContent += New-NumberedStep -text "Kartu sambutan di bagian atas menampilkan nama pengguna dan tombol Mulai Proyek untuk memulai proyek baru dengan cepat."
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\1-sudah-login-client.jpeg") -imgId $id1 -imgName "beranda-client" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Tiga kartu statistik menampilkan jumlah Proyek Aktif, Open Tender, dan Penawaran Masuk secara real-time."
$clientContent += New-NumberedStep -text "Bagian Menu Utama menyediakan empat akses cepat: Buat Proyek, Cari Kontraktor, Cari Arsitek, dan Lihat Progress."
$id2 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\2-menu-utama-beranda.jpeg") -imgId $id2 -imgName "menu-utama-beranda" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Bagian Mitra Terpopuler menampilkan daftar kontraktor dan arsitek dengan rating dan jumlah proyek terbanyak."
$clientContent += New-NumberedStep -text "Bagian Proyek Saya menampilkan proyek aktif beserta status dan persentase progress pembangunan terkini."
$clientContent += New-EmptyPara

# 2. Menu Mitra
$clientContent += New-HeadingL3 -text "Mencari dan Menghubungi Mitra (Kontraktor & Arsitek)"
$clientContent += New-TextPara -text "Client dapat menjelajahi dan menghubungi mitra (kontraktor atau arsitek) yang tersedia melalui menu Mitra. Sistem menyediakan dua tab terpisah untuk memudahkan pencarian berdasarkan peran."
$clientContent += New-BoldPara -text "Langkah-langkah mencari mitra:"

$id3 = Next-Id
$id4 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Mitra pada navigasi bawah. Halaman akan menampilkan dua tab: Kontraktor dan Arsitek."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\3-mitra1-kontraktor.jpeg") -imgId1 $id3 -imgName1 "mitra-kontraktor" -rId2 (Get-RId "2-client\3-mitra2-arsitek.jpeg") -imgId2 $id4 -imgName2 "mitra-arsitek" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Gunakan kolom pencarian untuk mencari nama kontraktor atau arsitek secara spesifik."
$clientContent += New-NumberedStep -text "Filter daftar menggunakan tombol sortir: Terbaru, Proyek Terbanyak, atau Rating Tertinggi."
$clientContent += New-NumberedStep -text "Ketuk tombol Lihat Profil untuk melihat portofolio, sertifikasi, dan riwayat proyek mitra."
$clientContent += New-NumberedStep -text "Ketuk tombol Konsultasi untuk langsung membuka sesi percakapan dengan mitra yang dipilih."
$clientContent += New-EmptyPara

# 3. Fitur Konsultasi
$clientContent += New-HeadingL3 -text "Fitur Konsultasi (Chat dengan Mitra)"
$clientContent += New-TextPara -text "Fitur Konsultasi memungkinkan Client berkomunikasi langsung dengan Kontraktor atau Arsitek sebelum memutuskan untuk menjalin kerja sama. Percakapan dilakukan melalui ruang obrolan interaktif yang terenkripsi dan tersimpan dalam riwayat."
$clientContent += New-BoldPara -text "Langkah-langkah memulai konsultasi:"

$id5 = Next-Id
$id6 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Konsultasi pada navigasi bawah atau tombol Konsultasi pada kartu mitra. Halaman ini menampilkan daftar percakapan aktif dengan label status (Pending/Aktif)."
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\4-konsultasi1.jpeg") -imgId $id5 -imgName "konsultasi-list" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Sebelum memulai konsultasi, konfirmasi minat Anda dengan menekan tombol yang tersedia pada halaman pra-konsultasi."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\4-pre-konsultasi1-kontraktor.jpeg") -imgId1 $id6 -imgName1 "pre-konsultasi-kontraktor" -rId2 (Get-RId "2-client\4-pre-konsultasi2-arsitek.jpeg") -imgId2 (Next-Id) -imgName2 "pre-konsultasi-arsitek" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk percakapan yang diinginkan untuk masuk ke ruang obrolan (Room Chat). Di sini Anda dapat mengirim pesan, gambar, dan dokumen terkait proyek."
$id7 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\4-roomchat-konsultasi1-kontraktor.jpeg") -imgId $id7 -imgName "roomchat-konsultasi" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

# 4. Progress Proyek
$clientContent += New-HeadingL3 -text "Memantau Progress Proyek"
$clientContent += New-TextPara -text "Menu Progress memungkinkan Client memantau perkembangan seluruh proyek yang sedang berjalan. Terdapat tiga tab utama: Pembangunan (progres fisik kontraktor), Desain (progres desain arsitek), dan Draft (dokumen rancangan awal)."
$clientContent += New-BoldPara -text "Panduan memantau progress pembangunan fisik:"

$id8 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Progress pada navigasi bawah. Sistem akan menampilkan halaman Proyek Saya dengan tab Pembangunan aktif secara default."
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\5-progress1-pembangunan1.jpeg") -imgId $id8 -imgName "progress-pembangunan" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk kartu proyek untuk melihat Detail Proyek yang mencakup informasi lokasi, nilai kontrak, dan deskripsi lengkap."
$id9 = Next-Id
$id10 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\5-progress1-pembangunan2-detail-proyek1.jpeg") -imgId1 $id9 -imgName1 "detail-proyek1" -rId2 (Get-RId "2-client\5-progress1-pembangunan2-detail-proyek2.jpeg") -imgId2 $id10 -imgName2 "detail-proyek2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Scroll ke bawah untuk melihat Detail Penawaran dari kontraktor, termasuk rincian RAB dan jadwal kerja."
$id11 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\5-progress-pembangunan2-detail-proyek3-detail-penawaran.jpeg") -imgId $id11 -imgName "detail-penawaran" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk bagian Termin Pembayaran untuk melihat jadwal dan status setiap termin pembayaran yang telah disepakati."
$id12 = Next-Id
$id13 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\5-progress1-pembangunan2-detail-proyek4-termin-pembayaran1.jpeg") -imgId1 $id12 -imgName1 "termin-pembayaran1" -rId2 (Get-RId "2-client\5-progress1-pembangunan2-detail-proyek4-termin-pembayaran2.jpeg") -imgId2 $id13 -imgName2 "termin-pembayaran2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

$clientContent += New-BoldPara -text "Panduan memantau progress desain arsitektur:"
$id14 = Next-Id
$id15 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk tab Desain pada halaman Progress untuk melihat proyek desain arsitektur yang sedang berjalan."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\5-progress2-desain1.jpeg") -imgId1 $id14 -imgName1 "progress-desain" -rId2 (Get-RId "2-client\5-progress2-desain1-detail-desain.jpeg") -imgId2 $id15 -imgName2 "detail-desain" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk kartu desain untuk melihat Detail Desain, termasuk file blueprint dan rancangan 3D yang diunggah oleh arsitek."
$clientContent += New-EmptyPara

$clientContent += New-BoldPara -text "Panduan memantau dan memperbarui Draft:"
$id16 = Next-Id
$id17 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk tab Draft untuk melihat dokumen rancangan awal yang masih dalam tahap revisi bersama arsitek."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\5-progress3-draft.jpeg") -imgId1 $id16 -imgName1 "progress-draft" -rId2 (Get-RId "2-client\6-progress-update-draft.jpeg") -imgId2 $id17 -imgName2 "update-draft" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Untuk memberikan masukan atau revisi pada draft, ketuk tombol yang tersedia dan tulis catatan revisi Anda. Perubahan akan langsung terkirim ke arsitek."
$clientContent += New-EmptyPara

# 5. Profil & Buat Proyek
$clientContent += New-HeadingL3 -text "Mengelola Profil dan Membuat Proyek Baru"
$clientContent += New-TextPara -text "Menu Profile menjadi pusat pengelolaan akun Client, termasuk melihat proyek aktif, mengedit data diri, dan membuat proyek konstruksi baru."
$clientContent += New-BoldPara -text "Mengedit data profil:"

$id18 = Next-Id
$id19 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Profile pada navigasi bawah untuk melihat ringkasan akun: nama, email, statistik proyek, dan daftar Proyek Saya."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\6-profile1.jpeg") -imgId1 $id18 -imgName1 "profile-client" -rId2 (Get-RId "2-client\6-profile3-update-profile.jpeg") -imgId2 $id19 -imgName2 "update-profile" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk ikon edit (pensil) di pojok kanan atas untuk masuk ke halaman Edit Profil. Perbarui foto, nama, nomor telepon, dan data lainnya, kemudian tekan Simpan."
$id20 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\6-profile1-edit-profile.jpeg") -imgId $id20 -imgName "edit-profile-client" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

$clientContent += New-BoldPara -text "Langkah-langkah membuat proyek baru:"
$id21 = Next-Id
$id22 = Next-Id
$clientContent += New-NumberedStep -text "Pada halaman Profile, ketuk tombol Proyek Saya lalu pilih Buat Proyek, atau langsung ketuk tombol Mulai Proyek / Buat Proyek Sekarang dari Beranda."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\6-profile2.jpeg") -imgId1 $id21 -imgName1 "profile2-client" -rId2 (Get-RId "2-client\6-profile2-buatproyek1.jpeg") -imgId2 $id22 -imgName2 "buat-proyek1" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Isi informasi dasar proyek: nama proyek, deskripsi, gaya arsitektur, luas bangunan, luas tanah, dan jumlah lantai."
$id23 = Next-Id
$id24 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\6-profile2-buatproyek2-1.jpeg") -imgId1 $id23 -imgName1 "buat-proyek2-1" -rId2 (Get-RId "2-client\6-profile2-buatproyek2-2.jpeg") -imgId2 $id24 -imgName2 "buat-proyek2-2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Isi informasi anggaran: estimasi biaya dan preferensi jangka waktu pembangunan."
$id25 = Next-Id
$id26 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\6-profile2-buatproyek3-1.jpeg") -imgId1 $id25 -imgName1 "buat-proyek3-1" -rId2 (Get-RId "2-client\6-profile2-buatproyek3-2.jpeg") -imgId2 $id26 -imgName2 "buat-proyek3-2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Unggah foto referensi atau dokumen pendukung proyek (opsional), kemudian tentukan lokasi proyek pada peta interaktif."
$id27 = Next-Id
$id28 = Next-Id
$id29 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\6-profile2-buatproyek4-1.jpeg") -imgId $id27 -imgName "buat-proyek4-1" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "2-client\6-profile2-buatproyek4-2.jpeg") -imgId1 $id28 -imgName1 "buat-proyek4-2" -rId2 (Get-RId "2-client\6-profile2-buatproyek4-3.jpeg") -imgId2 $id29 -imgName2 "buat-proyek4-3" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Tinjau kembali semua informasi yang telah diisi, lalu ketuk tombol Posting Proyek. Notifikasi konfirmasi akan muncul bahwa proyek berhasil dipublikasikan dan kini tersedia bagi mitra untuk mengajukan penawaran."
$id30 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "2-client\6-profile2-buatproyek5-notif.jpeg") -imgId $id30 -imgName "buat-proyek-notif" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

# ============================================================
# KONTEN BARU: KONTRAKTOR
# ============================================================
$clientContent += New-PageBreak
$clientContent += New-HeadingL2 -text "Panduan Pengguna Kontraktor (Mitra Pelaksana)" -bookmarkId (Next-Id) -bookmarkName "_Toc_Kontraktor01"

# 1. Beranda Kontraktor
$clientContent += New-HeadingL3 -text "Halaman Beranda Kontraktor"
$clientContent += New-TextPara -text "Setelah berhasil masuk ke akun, Kontraktor akan diarahkan ke halaman Beranda yang menampilkan ringkasan aktivitas bisnis, daftar penawaran masuk dari proyek-proyek aktif milik client, serta statistik kinerja secara keseluruhan."
$clientContent += New-BoldPara -text "Komponen utama halaman Beranda Kontraktor:"

$id_k1 = Next-Id
$clientContent += New-NumberedStep -text "Kartu profil di bagian atas menampilkan nama, perusahaan, dan tombol Lengkapi Portofolio / Sertifikasi untuk memperkuat kredibilitas profil."
$clientContent += New-ImageParagraph -rId (Get-RId "3-kontraktor\1-sudah-login-client.jpeg") -imgId $id_k1 -imgName "beranda-kontraktor" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Tiga kartu statistik menampilkan jumlah Proyek Aktif, Selesai, dan Rating bintang rata-rata dari klien."
$clientContent += New-NumberedStep -text "Bagian Penawaran Masuk menampilkan proyek-proyek baru yang dipublikasikan oleh client dan sedang membuka tender untuk kontraktor. Ketuk Lihat Semua untuk melihat daftar lengkap."
$clientContent += New-EmptyPara

# 2. Penawaran Proyek
$clientContent += New-HeadingL3 -text "Melihat dan Mengajukan Penawaran Proyek"
$clientContent += New-TextPara -text "Kontraktor dapat melihat detail proyek yang terbuka dan mengajukan Rencana Anggaran Biaya (RAB) sebagai dokumen penawaran resmi. Proses ini dilakukan melalui menu Proyek pada navigasi bawah."
$clientContent += New-BoldPara -text "Langkah-langkah mengajukan penawaran:"

$id_k2 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Proyek pada navigasi bawah untuk melihat daftar seluruh proyek yang sedang membuka penawaran."
$clientContent += New-ImageParagraph -rId (Get-RId "3-kontraktor\2-penawaran-masuk1-detail-proyek1.jpeg") -imgId $id_k2 -imgName "penawaran-masuk-detail" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk kartu proyek untuk melihat halaman Detail Proyek yang mencakup deskripsi lengkap: luas bangunan, luas tanah, gaya arsitektur, jumlah lantai, dan estimasi anggaran client."
$clientContent += New-NumberedStep -text "Jika proyek sesuai dengan kapasitas, ketuk tombol Ajukan Penawaran. Isi formulir penawaran dengan total nilai RAB, catatan teknis, dan jadwal pengerjaan yang diusulkan."
$id_k3 = Next-Id
$id_k4 = Next-Id
$id_k5 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "3-kontraktor\2-penawaran-masuk2-diajukan1.jpeg") -imgId $id_k3 -imgName "penawaran-diajukan1" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "3-kontraktor\2-penawaran-masuk2-diajukan2.jpeg") -imgId1 $id_k4 -imgName1 "penawaran-diajukan2" -rId2 (Get-RId "3-kontraktor\2-penawaran-masuk2-diajukan3.jpeg") -imgId2 $id_k5 -imgName2 "penawaran-diajukan3" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Setelah formulir terisi lengkap, ketuk Kirim Penawaran. Status penawaran akan berubah menjadi Menunggu Keputusan hingga client menerima atau menolak tawaran tersebut."
$clientContent += New-EmptyPara

# 3. Progress Kontraktor
$clientContent += New-HeadingL3 -text "Melaporkan dan Memantau Progress Pembangunan"
$clientContent += New-TextPara -text "Menu Progress memungkinkan Kontraktor memantau proyek yang sedang dikerjakan dan secara aktif melaporkan kemajuan pembangunan secara berkala kepada client. Laporan progress yang rutin meningkatkan kepercayaan dan transparansi."
$clientContent += New-BoldPara -text "Langkah-langkah melaporkan kemajuan proyek:"

$id_k6 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Progress pada navigasi bawah untuk melihat daftar proyek aktif yang sedang dikerjakan."
$clientContent += New-ImageParagraph -rId (Get-RId "3-kontraktor\3-progress1.jpeg") -imgId $id_k6 -imgName "progress-kontraktor" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk proyek aktif untuk masuk ke halaman Detail Penawaran. Di sini tersedia rincian nilai kontrak yang disepakati, catatan pekerjaan, dan status termin pembayaran."
$id_k7 = Next-Id
$id_k8 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "3-kontraktor\3-progress1-detail-penawaran1.jpeg") -imgId1 $id_k7 -imgName1 "detail-penawaran-k1" -rId2 (Get-RId "3-kontraktor\3-progress1-detail-penawaran2.jpeg") -imgId2 $id_k8 -imgName2 "detail-penawaran-k2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk bagian Termin Pembayaran untuk melihat jadwal dan status pembayaran dari client. Kontraktor dapat memperbarui status pekerjaan sesuai dengan termin yang telah selesai."
$id_k9 = Next-Id
$id_k10 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "3-kontraktor\3-progress1-detail-penawaran3-termin-pembayaran1.jpeg") -imgId1 $id_k9 -imgName1 "termin-kontraktor1" -rId2 (Get-RId "3-kontraktor\3-progress1-detail-penawaran3-termin-pembayaran2.jpeg") -imgId2 $id_k10 -imgName2 "termin-kontraktor2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

# 4. Edit Profil Kontraktor
$clientContent += New-HeadingL3 -text "Mengelola Profil dan Portofolio Kontraktor"
$clientContent += New-TextPara -text "Profil yang lengkap dan terverifikasi meningkatkan peluang kontraktor untuk mendapatkan proyek. Menu Profile memungkinkan kontraktor memperbarui data perusahaan, portofolio, dan sertifikasi profesional."
$clientContent += New-BoldPara -text "Langkah-langkah mengedit profil kontraktor:"

$id_k11 = Next-Id
$id_k12 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Profile pada navigasi bawah. Halaman ini menampilkan informasi akun, statistik proyek, dan tombol Edit Profil."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "3-kontraktor\4-edit-profil1.jpeg") -imgId1 $id_k11 -imgName1 "profil-kontraktor1" -rId2 (Get-RId "3-kontraktor\4-edit-profil2.jpeg") -imgId2 $id_k12 -imgName2 "profil-kontraktor2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk Edit Profil untuk masuk ke formulir pengeditan. Perbarui nama, nama perusahaan, nomor telepon, deskripsi layanan, dan foto profil."
$id_k13 = Next-Id
$id_k14 = Next-Id
$id_k15 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "3-kontraktor\4-edit-profil3-1.jpeg") -imgId $id_k13 -imgName "edit-profil-k1" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "3-kontraktor\4-edit-profil3-2.jpeg") -imgId1 $id_k14 -imgName1 "edit-profil-k2" -rId2 (Get-RId "3-kontraktor\4-edit-profil3-3.jpeg") -imgId2 $id_k15 -imgName2 "edit-profil-k3" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Tambahkan dokumen sertifikasi profesional (SKA/SKT) dan foto portofolio proyek sebelumnya untuk meningkatkan kepercayaan client. Setelah semua data diisi, ketuk Simpan. Notifikasi konfirmasi akan muncul sebagai tanda profil berhasil diperbarui."
$id_k16 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "3-kontraktor\4-edit-profil4-notif.jpeg") -imgId $id_k16 -imgName "edit-profil-notif" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

# ============================================================
# KONTEN BARU: ARSITEK
# ============================================================
$clientContent += New-PageBreak
$clientContent += New-HeadingL2 -text "Panduan Pengguna Arsitek (Mitra Perencana)" -bookmarkId (Next-Id) -bookmarkName "_Toc_Arsitek01"

# 1. Beranda Arsitek
$clientContent += New-HeadingL3 -text "Halaman Beranda Arsitek"
$clientContent += New-TextPara -text "Setelah berhasil masuk, Arsitek akan diarahkan ke halaman Beranda yang dirancang khusus untuk kebutuhan mitra perencana. Halaman ini menampilkan progres kelengkapan profil, statistik kolaborasi aktif, total desain yang diunggah, serta galeri desain populer."
$clientContent += New-BoldPara -text "Komponen utama halaman Beranda Arsitek:"

$id_a1 = Next-Id
$clientContent += New-NumberedStep -text "Banner profil di bagian atas menampilkan nama arsitek dan persentase Kelengkapan Profil beserta progress bar. Semakin lengkap profil, semakin tinggi peluang menjangkau klien."
$clientContent += New-ImageParagraph -rId (Get-RId "4-arsitek\1-sudah-login-arsitek1.jpeg") -imgId $id_a1 -imgName "beranda-arsitek" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Empat kartu statistik menampilkan: Kolaborasi Aktif (jumlah proyek yang sedang dikerjakan bersama client), Total Desain (jumlah file desain yang telah diunggah), Tahun Pengalaman, dan jumlah Sertifikasi."
$clientContent += New-NumberedStep -text "Bagian Desain Populer di bawah menampilkan galeri karya arsitektur yang paling banyak dilihat oleh pengguna."
$clientContent += New-NumberedStep -text "Navigasi bawah terdiri dari empat menu utama: Beranda, Desain, Inbox, dan Profil."
$clientContent += New-EmptyPara

# 2. Mengelola Desain
$clientContent += New-HeadingL3 -text "Mengelola dan Mengunggah Desain"
$clientContent += New-TextPara -text "Menu Desain adalah area kerja utama arsitek untuk mengunggah, mengelola, dan mempublikasikan karya desain arsitektur. Desain yang diunggah akan dapat diakses oleh client yang menggunakan layanan arsitek tersebut."
$clientContent += New-BoldPara -text "Langkah-langkah mengunggah desain baru:"

$id_a2 = Next-Id
$id_a3 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Desain pada navigasi bawah. Halaman ini menampilkan galeri seluruh karya desain yang telah diunggah beserta tombol tambah desain baru."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "4-arsitek\2-desain1-upload-desain1.jpeg") -imgId1 $id_a2 -imgName1 "upload-desain1" -rId2 (Get-RId "4-arsitek\2-desain1-upload-desain2.jpeg") -imgId2 $id_a3 -imgName2 "upload-desain2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk tombol tambah (+) atau Unggah Desain. Pilih file desain dari penyimpanan perangkat (format yang didukung: PNG, JPG, PDF)."
$clientContent += New-NumberedStep -text "Isi judul desain, deskripsi singkat, dan kategori (misalnya: Minimalis Modern, Tropis Kontemporer). Tambahkan tag untuk memudahkan pencarian."
$clientContent += New-NumberedStep -text "Ketuk Publikasikan untuk menyimpan dan mempublikasikan desain ke galeri BuildMatch. Desain yang dipublikasikan dapat dilihat oleh seluruh client yang menggunakan platform."
$clientContent += New-EmptyPara

# 3. Inbox / Chat
$clientContent += New-HeadingL3 -text "Mengelola Percakapan Klien (Inbox)"
$clientContent += New-TextPara -text "Menu Inbox adalah pusat komunikasi arsitek dengan para klien. Arsitek dapat menerima permintaan konsultasi, membalas pesan, dan mengirimkan file desain langsung melalui ruang obrolan yang tersedia."
$clientContent += New-BoldPara -text "Langkah-langkah menggunakan fitur Inbox:"

$id_a4 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Inbox pada navigasi bawah. Halaman ini menampilkan daftar seluruh percakapan dengan klien. Pesan yang belum dibaca ditandai dengan indikator biru."
$clientContent += New-ImageParagraph -rId (Get-RId "4-arsitek\3-inbox1-belum-terbaca.jpeg") -imgId $id_a4 -imgName "inbox-belum-terbaca" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk percakapan untuk membuka ruang obrolan (Room Chat). Arsitek dapat membaca pesan klien, membalas, dan mengirimkan lampiran file desain."
$id_a5 = Next-Id
$id_a6 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "4-arsitek\3-inbox2-roomchat1-1.jpeg") -imgId1 $id_a5 -imgName1 "roomchat-arsitek1" -rId2 (Get-RId "4-arsitek\3-inbox2-roomchat1-2.jpeg") -imgId2 $id_a6 -imgName2 "roomchat-arsitek2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Setelah membaca dan membalas pesan, status percakapan akan berubah menjadi sudah terbaca. Daftar Inbox akan diperbarui secara otomatis."
$id_a7 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "4-arsitek\3-inbox3-sudah-terbaca.jpeg") -imgId $id_a7 -imgName "inbox-sudah-terbaca" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

# 4. Profil Arsitek
$clientContent += New-HeadingL3 -text "Mengelola Profil Arsitek"
$clientContent += New-TextPara -text "Profil arsitek yang komprehensif dan menarik adalah kunci untuk mendapatkan kepercayaan dari calon klien. Menu Profil memungkinkan arsitek untuk memperbarui informasi pribadi, portofolio, dan sertifikasi profesional."
$clientContent += New-BoldPara -text "Langkah-langkah melihat dan mengedit profil arsitek:"

$id_a8 = Next-Id
$id_a9 = Next-Id
$clientContent += New-NumberedStep -text "Ketuk menu Profil pada navigasi bawah. Halaman profil menampilkan foto, nama, spesialisasi, dan ringkasan statistik arsitek."
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "4-arsitek\4-profil1.jpeg") -imgId1 $id_a8 -imgName1 "profil-arsitek1" -rId2 (Get-RId "4-arsitek\4-profil2.jpeg") -imgId2 $id_a9 -imgName2 "profil-arsitek2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Scroll ke bawah untuk melihat bagian portofolio karya, sertifikasi profesional yang telah diunggah, dan ulasan dari klien sebelumnya."
$id_a10 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "4-arsitek\4-profil3.jpeg") -imgId $id_a10 -imgName "profil-arsitek3" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Ketuk tombol Edit Profil untuk masuk ke formulir pengeditan. Perbarui foto profil, nama lengkap, deskripsi keahlian, nomor telepon, dan spesialisasi desain."
$id_a11 = Next-Id
$id_a12 = Next-Id
$clientContent += New-TwoImageParagraph -rId1 (Get-RId "4-arsitek\4-profil5-edit-profil1.jpeg") -imgId1 $id_a11 -imgName1 "edit-profil-a1" -rId2 (Get-RId "4-arsitek\4-profil5-edit-profl2.jpeg") -imgId2 $id_a12 -imgName2 "edit-profil-a2" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-NumberedStep -text "Tambahkan atau perbarui dokumen sertifikasi (IAI, ASEAN Architect, dll.) dan file portofolio karya terbaru. Setelah semua informasi diisi dengan benar, ketuk tombol Simpan untuk menyimpan perubahan."
$id_a13 = Next-Id
$clientContent += New-ImageParagraph -rId (Get-RId "4-arsitek\4-profil5-edit-profil3.jpeg") -imgId $id_a13 -imgName "edit-profil-a3" -widthEmu 1346200 -heightEmu 2979040
$clientContent += New-EmptyPara

Write-Host "Konten XML selesai dibuat ($($clientContent.Length) karakter)" -ForegroundColor Green

# Baca document.xml yang ada
$docXmlFile = "$tempDir\docx\word\document.xml"
$docXml = Get-Content $docXmlFile -Raw

# Cari posisi untuk menyisipkan konten baru (sebelum tag </w:body>)
$insertPoint = $docXml.LastIndexOf("</w:body>")
if ($insertPoint -lt 0) {
    Write-Error "Tidak dapat menemukan tag </w:body> dalam document.xml!"
    exit 1
}

# Sisipkan konten baru
$newDocXml = $docXml.Substring(0, $insertPoint) + $clientContent + $docXml.Substring($insertPoint)
Set-Content $docXmlFile $newDocXml -Encoding UTF8

Write-Host "document.xml berhasil diperbarui" -ForegroundColor Green

# Buat ulang DOCX dari folder yang sudah dimodifikasi
$outputZip = "$tempDir\output.zip"
if (Test-Path $outputZip) { Remove-Item $outputZip }

# Compress folder docx menjadi zip
Add-Type -Assembly 'System.IO.Compression.FileSystem'
[System.IO.Compression.ZipFile]::CreateFromDirectory("$tempDir\docx", $outputZip)

# Rename zip menjadi docx
if (Test-Path $outputDocx) { Remove-Item $outputDocx }
Copy-Item $outputZip $outputDocx

Write-Host "=== File output berhasil dibuat: $outputDocx ===" -ForegroundColor Green
Write-Host "Ukuran file: $([Math]::Round((Get-Item $outputDocx).Length / 1MB, 2)) MB" -ForegroundColor Cyan

# Cleanup
Remove-Item "$tempDir\source.zip"
# Remove-Item $tempDir -Recurse -Force  # Biarkan untuk debugging

Write-Host "Proses selesai! Silakan buka file '$outputDocx'" -ForegroundColor Yellow
