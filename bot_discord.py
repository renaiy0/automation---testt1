"""
Bot Discord: kirim file .xlsx di channel/DM -> dibalas file hasil + total.

Setup:
  1. https://discord.com/developers/applications -> New Application -> Bot -> Reset Token (salin).
     Di halaman Bot, aktifkan "Message Content Intent".
  2. OAuth2 -> URL Generator: scope "bot", izin "Send Messages", "Attach Files",
     "Read Message History" -> buka URL-nya untuk mengundang bot ke servermu.
  3. pip install discord.py pandas openpyxl
  4. Windows: set DISCORD_TOKEN=isi_token   (Linux/Mac: export DISCORD_TOKEN=isi_token)
     (opsional) set ALLOWED_IDS=id_user_discord_kamu
  5. python bot_discord.py
Caption (teks yang ditulis bersama file) opsional: qty=NamaKolomJumlah harga=NamaKolomHarga
"""
import asyncio
import os
import tempfile
from pathlib import Path

import discord

from bot_common import hitung

TOKEN = os.environ["BSHLeRMFFrqY7OEvm3QOJ3Qyt5HEqHDt"]
ALLOWED = {int(x) for x in os.environ.get("ALLOWED_IDS", "").split(",") if x.strip()}

intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)


@client.event
async def on_ready():
    print(f"Bot RUNNING {client.user}")


@client.event
async def on_message(msg: discord.Message):
    if msg.author.bot or (ALLOWED and msg.author.id not in ALLOWED):
        return
    for att in msg.attachments:
        if not att.filename.lower().endswith(".xlsx"):
            continue
        await msg.channel.send("Diproses...")
        tmp = Path(tempfile.mkdtemp(prefix="masuk_")) / att.filename
        await att.save(tmp)
        try:
            # proses sinkron -> jalankan di thread supaya bot tidak nge-freeze
            hasil, teks = await asyncio.to_thread(hitung, tmp, msg.content)
        except ValueError as e:
            await msg.channel.send(f"Gagal: {e}")
            continue
        except Exception as e:
            await msg.channel.send(f"File tidak bisa dibaca: {e}")
            continue
        await msg.channel.send(teks, file=discord.File(hasil))


client.run(TOKEN)
