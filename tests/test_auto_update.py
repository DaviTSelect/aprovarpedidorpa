import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import auto_update


class BuscarReleaseTests(unittest.TestCase):
    @patch("requests.get")
    def test_busca_release_mais_recente(self, get):
        response = Mock()
        response.json.return_value = {"tag_name": "v1.0.2"}
        get.return_value = response

        release = auto_update.buscar_release("DaviTSelect/aprovarpedidorpa")

        self.assertEqual(release, {"tag_name": "v1.0.2"})
        get.assert_called_once_with(
            "https://api.github.com/repos/DaviTSelect/aprovarpedidorpa/releases/latest",
            timeout=15,
            headers={"Accept": "application/vnd.github+json"},
        )
        response.raise_for_status.assert_called_once_with()
        response.json.assert_called_once_with()

    @patch("requests.get")
    def test_propaga_erro_http_da_api(self, get):
        response = Mock()
        response.raise_for_status.side_effect = RuntimeError("HTTP 503")
        get.return_value = response

        with self.assertRaisesRegex(RuntimeError, "HTTP 503"):
            auto_update.buscar_release("org/repo")


class CompararVersoesTests(unittest.TestCase):
    def test_identifica_versoes_mais_novas_com_prefixo_opcional(self):
        casos = [
            ({"tag_name": "v1.0.2"}, "1.0.1", True),
            ({"tag_name": "1.1.0"}, "v1.0.9", True),
            ({"tag_name": "2.0.0"}, "1.99.99", True),
            ({"tag_name": "v1.0.1"}, "1.0.1", False),
            ({"tag_name": "v1.0.0"}, "1.0.1", False),
            ({"tag_name": "v1.0.9"}, "1.0.10", False),
        ]
        for release, atual, esperado in casos:
            with self.subTest(release=release, atual=atual):
                self.assertEqual(auto_update.tem_versao_nova(release, atual), esperado)

    def test_release_sem_tag_nao_e_mais_nova(self):
        self.assertFalse(auto_update.tem_versao_nova({}, "1.0.1"))


class PodeInstalarTests(unittest.TestCase):
    def test_so_permite_instalacao_no_executavel_compilado(self):
        with patch.object(sys, "frozen", True, create=True):
            self.assertTrue(auto_update.pode_instalar())

        with patch.object(sys, "frozen", False, create=True):
            self.assertFalse(auto_update.pode_instalar())


class InstalarTests(unittest.TestCase):
    def test_asset_ausente_falha_sem_tentar_download(self):
        with patch("requests.get") as get:
            with self.assertRaisesRegex(RuntimeError, "AprovarPedidosPF.exe"):
                auto_update.instalar(
                    {"tag_name": "v1.0.2", "assets": [{"name": "outro.exe"}]},
                    "AprovarPedidosPF.exe",
                )
        get.assert_not_called()

    def test_download_grava_arquivo_agenda_substituicao_e_reinicio(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            exe_atual = Path(temp_dir) / "AprovarPedidosPF.exe"
            exe_novo = Path(temp_dir) / "AprovarPedidosPF.novo.exe"
            arquivo_batch = Path(temp_dir) / "_atualizar_aprovar.bat"
            release = {
                "tag_name": "v1.0.2",
                "assets": [{
                    "name": "AprovarPedidosPF.exe",
                    "browser_download_url": "https://example.test/AprovarPedidosPF.exe",
                }],
            }
            response = Mock()
            response.iter_content.return_value = [b"novo-", b"", b"executavel"]

            with (
                patch.object(sys, "executable", str(exe_atual)),
                patch("requests.get", return_value=response) as get,
                patch("auto_update.subprocess.Popen") as popen,
            ):
                auto_update.instalar(release, "AprovarPedidosPF.exe")

            get.assert_called_once_with(
                "https://example.test/AprovarPedidosPF.exe",
                stream=True,
                timeout=180,
            )
            response.raise_for_status.assert_called_once_with()
            response.iter_content.assert_called_once_with(chunk_size=65536)
            self.assertEqual(exe_novo.read_bytes(), b"novo-executavel")

            conteudo_batch = arquivo_batch.read_text(encoding="utf-8")
            self.assertIn(f'move /y "{exe_novo}" "{exe_atual}"', conteudo_batch)
            self.assertIn(f'start "" "{exe_atual}"', conteudo_batch)
            self.assertIn('del "%~f0"', conteudo_batch)
            popen.assert_called_once_with(
                ["cmd", "/c", str(arquivo_batch)],
                creationflags=(
                    subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS
                ),
            )

    def test_propaga_erro_http_do_download(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            exe_atual = Path(temp_dir) / "AprovarPedidosPF.exe"
            release = {
                "tag_name": "v1.0.2",
                "assets": [{
                    "name": "AprovarPedidosPF.exe",
                    "browser_download_url": "https://example.test/AprovarPedidosPF.exe",
                }],
            }
            response = Mock()
            response.raise_for_status.side_effect = RuntimeError("download indisponivel")

            with (
                patch.object(sys, "executable", str(exe_atual)),
                patch("requests.get", return_value=response),
            ):
                with self.assertRaisesRegex(RuntimeError, "download indisponivel"):
                    auto_update.instalar(release, "AprovarPedidosPF.exe")

            self.assertFalse((Path(temp_dir) / "AprovarPedidosPF.novo.exe").exists())
            self.assertFalse((Path(temp_dir) / "_atualizar_aprovar.bat").exists())


class FluxoIntegradoAtualizacaoTests(unittest.TestCase):
    def test_release_nova_e_baixada_e_preparada_para_instalacao(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            exe_atual = Path(temp_dir) / "AprovarPedidosPF.exe"
            release_response = Mock()
            release_response.json.return_value = {
                "tag_name": "v1.0.2",
                "assets": [{
                    "name": "AprovarPedidosPF.exe",
                    "browser_download_url": "https://example.test/AprovarPedidosPF.exe",
                }],
            }
            download_response = Mock()
            download_response.iter_content.return_value = [b"novo-executavel"]

            with (
                patch.object(sys, "executable", str(exe_atual)),
                patch(
                    "requests.get",
                    side_effect=[release_response, download_response],
                ) as get,
                patch("auto_update.subprocess.Popen") as popen,
            ):
                release = auto_update.buscar_release("DaviTSelect/aprovarpedidorpa")
                deve_atualizar = auto_update.tem_versao_nova(release, "1.0.1")
                if deve_atualizar:
                    auto_update.instalar(release, auto_update.NOME_EXE)

            self.assertTrue(deve_atualizar)
            self.assertEqual(
                (Path(temp_dir) / "AprovarPedidosPF.novo.exe").read_bytes(),
                b"novo-executavel",
            )
            self.assertEqual(get.call_count, 2)
            popen.assert_called_once()


if __name__ == "__main__":
    unittest.main()
