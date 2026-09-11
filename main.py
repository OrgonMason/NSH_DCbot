import discord
from discord.ext import commands, tasks
from discord import app_commands
import json
import os
from keep_alive import keep_alive
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from typing import Dict, List, Tuple, Optional
import asyncio
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from collections import defaultdict
import itertools

load_dotenv()

# ==================== 設定區 ====================
TOKEN = os.getenv("DISCORD_TOKEN")

try:
    TZ = ZoneInfo("Asia/Taipei")
except Exception:
    # 如果沒有 tzdata，就用固定 UTC+8
    from datetime import timezone, timedelta
    TZ = timezone(timedelta(hours=8))

# 九職業（劍網三）
JOBS = [
    "鐵衣", "血河", "九靈", "素問", "神相",
    "龍吟", "碎夢", "玄機", "潮光"
]

MAX_PLAYERS = 60
MAX_PER_JOB = 14
NUM_GROUPS = 10
PLAYERS_PER_GROUP = 6

DATA_FILE = "event_data.json"
WEEKDAY_ZH = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
# ================================================

intents = discord.Intents.default()
intents.message_content = True
intents.members = True
intents.guilds = True

bot = commands.Bot(command_prefix="!", intents=intents)


# ---------- 資料處理 ----------
def load_data() -> dict:
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_data(data: dict):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_guild_data(guild_id: int) -> dict:
    data = load_data()
    gid = str(guild_id)
    if gid not in data:
        data[gid] = {
            "event_name": None,
            "event_date": None,
            "opponent": None,
            "deadline": None,
            "deadline_iso": None,
            "channel_id": None,
            "thread_id": None,
            "message_id": None,
            "signups": {},
            "registration_closed": False
        }
        save_data(data)
    return data[gid]


def update_guild_data(guild_id: int, guild_data: dict):
    data = load_data()
    data[str(guild_id)] = guild_data
    save_data(data)


# ---------- 時間工具 ----------
def format_datetime_display(dt: datetime) -> str:
    weekday = WEEKDAY_ZH[dt.weekday()]
    return f"{dt.strftime('%Y-%m-%d')}【{weekday}】 {dt.strftime('%H:%M')}"


def parse_and_build_datetime(date_str: str, hour: int, minute: int) -> Optional[datetime]:
    try:
        year, month, day = map(int, date_str.strip().split("-"))
        dt = datetime(year, month, day, hour, minute, tzinfo=TZ)
        return dt
    except Exception:
        return None


def is_deadline_passed(guild_data: dict) -> bool:
    if guild_data.get("registration_closed"):
        return True
    deadline_iso = guild_data.get("deadline_iso")
    if not deadline_iso:
        return False
    try:
        deadline = datetime.fromisoformat(deadline_iso)
        now = datetime.now(TZ)
        return now >= deadline
    except Exception:
        return False


# ---------- 狀態計算 ----------
def get_job_counts(signups: dict) -> Dict[str, int]:
    counts = {job: 0 for job in JOBS}
    for info in signups.values():
        job = info.get("job")
        if job in counts:
            counts[job] += 1
    return counts


def get_remaining_slots(signups: dict) -> int:
    return MAX_PLAYERS - len(signups)


def format_status_embed(guild_data: dict) -> discord.Embed:
    signups = guild_data.get("signups", {})
    job_counts = get_job_counts(signups)
    remaining = get_remaining_slots(signups)
    total = len(signups)
    closed = is_deadline_passed(guild_data)

    color = discord.Color.red() if closed or remaining <= 0 else discord.Color.blue()

    embed = discord.Embed(
        title=f"⚔️ {guild_data.get('event_name', '聯賽活動')} 報名狀態",
        color=color
    )

    embed.add_field(name="🏰 對手幫會", value=guild_data.get("opponent") or "未設定", inline=True)
    embed.add_field(name="⏰ 約戰時間", value=guild_data.get("event_date") or "未設定", inline=True)
    embed.add_field(name="📅 報名截止", value=guild_data.get("deadline") or "未設定", inline=True)

    status_text = "🔴 **報名已截止**" if closed else f"🟢 報名進行中（剩餘 **{remaining}** 名額）"
    embed.add_field(
        name="📊 目前人數",
        value=f"{total} / {MAX_PLAYERS}\n{status_text}",
        inline=False
    )

    job_lines = []
    for job in JOBS:
        count = job_counts[job]
        if count >= MAX_PER_JOB:
            status = "🔴 已滿"
        else:
            status = f"🟢 剩餘 {MAX_PER_JOB - count}"
        job_lines.append(f"**{job}**：{count}/{MAX_PER_JOB}　{status}")
    embed.add_field(name="🗡️ 職業報名狀況（每職業上限 14 人）", value="\n".join(job_lines), inline=False)

    if signups:
        by_job = {job: [] for job in JOBS}
        for info in signups.values():
            job = info.get("job")
            char_name = info.get("char_name") or "未知"
            if job in by_job:
                by_job[job].append(f"{char_name}【{job}】")

        list_lines = []
        for job in JOBS:
            if by_job[job]:
                list_lines.append(f"**{job}**\n" + "、".join(by_job[job]))

        full_list = "\n\n".join(list_lines)
        if len(full_list) > 1000:
            full_list = full_list[:980] + "\n…（名單過長已截斷）"

        embed.add_field(name="👥 已報名成功人員", value=full_list, inline=False)
    else:
        embed.add_field(name="👥 已報名成功人員", value="目前尚無人報名", inline=False)

    if closed:
        embed.set_footer(text="⛔ 報名已截止，無法再進行報名或取消")
    else:
        embed.set_footer(text="請使用下方選單選擇職業報名，或按按鈕取消報名")
    embed.timestamp = datetime.now(TZ)
    return embed


async def update_thread_status(guild: discord.Guild):
    guild_data = get_guild_data(guild.id)
    thread_id = guild_data.get("thread_id")
    message_id = guild_data.get("message_id")

    if not thread_id or not message_id:
        return

    try:
        thread = guild.get_thread(thread_id) or await bot.fetch_channel(thread_id)
    except Exception:
        return

    closed = is_deadline_passed(guild_data)
    if closed and not guild_data.get("registration_closed"):
        guild_data["registration_closed"] = True
        update_guild_data(guild.id, guild_data)

    embed = format_status_embed(guild_data)
    view = SignupView(disabled=closed)

    try:
        msg = await thread.fetch_message(message_id)
        await msg.edit(embed=embed, view=view)
    except Exception:
        new_msg = await thread.send(embed=embed, view=view)
        guild_data["message_id"] = new_msg.id
        update_guild_data(guild.id, guild_data)


# ---------- 智能分組演算法 ----------
def auto_group_players(signups: dict) -> List[List[dict]]:
    """
    將報名玩家分成 10 組，每組最多 6 人。
    規則：
    1. 每組職業盡量不重複
    2. 人數平均分配
    3. 若某職業超過 10 人，多出來的人分配到人數較少的組
    """
    # 準備玩家列表（含 user_id）
    players = []
    for uid, info in signups.items():
        players.append({
            "uid": uid,
            "char_name": info.get("char_name", "未知"),
            "job": info.get("job", ""),
            "name": info.get("name", ""),
            "joined_at": info.get("joined_at", "")
        })

    # 按職業分桶
    job_buckets: Dict[str, List[dict]] = defaultdict(list)
    for p in players:
        job_buckets[p["job"]].append(p)

    # 初始化 10 個空組
    groups: List[List[dict]] = [[] for _ in range(NUM_GROUPS)]
    group_jobs: List[set] = [set() for _ in range(NUM_GROUPS)]  # 每組已有的職業

    # 第一輪：每個職業先盡量一人一組（最多 10 人）
    for job in JOBS:
        people = job_buckets.get(job, [])
        # 先分配前 10 人，一人一組
        for i, person in enumerate(people[:NUM_GROUPS]):
            groups[i].append(person)
            group_jobs[i].add(job)

        # 超過 10 人的，先放回剩餘
        job_buckets[job] = people[NUM_GROUPS:]

    # 第二輪：把剩餘的人（職業重複的）分配到人數最少、且該組還沒有這個職業的組
    remaining = []
    for job, people in job_buckets.items():
        remaining.extend(people)

    # 若還有人，就分配到人數最少的組（優先選沒有該職業的組）
    for person in remaining:
        job = person["job"]
        # 找出人數最少的組
        candidates = []
        for i in range(NUM_GROUPS):
            if len(groups[i]) >= PLAYERS_PER_GROUP:
                continue
            has_job = job in group_jobs[i]
            candidates.append((len(groups[i]), 1 if has_job else 0, i))

        if not candidates:
            # 所有組都滿了，硬塞到第一組（理論上不會）
            groups[0].append(person)
            continue

        # 優先：人數少 → 沒有該職業
        candidates.sort()
        best_idx = candidates[0][2]
        groups[best_idx].append(person)
        group_jobs[best_idx].add(job)

    # 最後再平衡一下（把人數差太大的微調，但保持職業不重複為優先）
    # 簡單處理：已經夠用了
    return groups


# ---------- Modal ----------
class CharacterNameModal(discord.ui.Modal, title="輸入角色名稱"):
    def __init__(self, job: str):
        super().__init__()
        self.job = job
        self.char_name_input = discord.ui.TextInput(
            label="你的角色名稱",
            placeholder="請輸入遊戲內的角色名稱",
            min_length=1,
            max_length=20,
            required=True
        )
        self.add_item(self.char_name_input)

    async def on_submit(self, interaction: discord.Interaction):
        char_name = self.char_name_input.value.strip()
        await complete_signup(interaction, self.job, char_name)


# ---------- 報名完成 ----------
async def complete_signup(interaction: discord.Interaction, job: str, char_name: str):
    guild = interaction.guild
    if guild is None:
        await interaction.response.send_message("❌ 此功能只能在伺服器內使用。", ephemeral=True)
        return

    guild_data = get_guild_data(guild.id)

    if is_deadline_passed(guild_data):
        await interaction.response.send_message("⛔ 報名已截止，無法再報名。", ephemeral=True)
        await asyncio.sleep(5)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass
        return

    if not guild_data.get("event_name"):
        await interaction.response.send_message("❌ 目前沒有進行中的活動。", ephemeral=True)
        return

    user_id = str(interaction.user.id)

    if user_id in guild_data["signups"]:
        old = guild_data["signups"][user_id]
        await interaction.response.send_message(
            f"⚠️ 你已經報名過了！\n角色：**{old.get('char_name', '未知')}**　職業：**{old['job']}**\n"
            f"如需更改請先按「取消報名」。",
            ephemeral=True
        )
        await asyncio.sleep(6)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass
        return

    if len(guild_data["signups"]) >= MAX_PLAYERS:
        await interaction.response.send_message("❌ 報名人數已達總上限（60 人）！", ephemeral=True)
        await asyncio.sleep(5)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass
        return

    job_counts = get_job_counts(guild_data["signups"])
    if job_counts.get(job, 0) >= MAX_PER_JOB:
        await interaction.response.send_message(
            f"❌ **{job}** 已達上限（{MAX_PER_JOB} 人），請選擇其他職業。",
            ephemeral=True
        )
        await asyncio.sleep(5)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass
        return

    guild_data["signups"][user_id] = {
        "job": job,
        "char_name": char_name,
        "name": str(interaction.user),
        "joined_at": datetime.now(TZ).isoformat()
    }
    update_guild_data(guild.id, guild_data)
    await update_thread_status(guild)

    remaining = get_remaining_slots(guild_data["signups"])

    # 取得討論串連結
    thread_link = ""
    thread_id = guild_data.get("thread_id")
    if thread_id:
        thread_link = f"https://discord.com/channels/{guild.id}/{thread_id}"

    await interaction.response.send_message(
        f"✅ **報名成功！**\n"
        f"活動：{guild_data['event_name']}\n"
        f"對手：{guild_data.get('opponent', '未設定')}\n"
        f"約戰時間：{guild_data.get('event_date', '未設定')}\n"
        f"角色名稱：**{char_name}**\n"
        f"職業：**{job}**\n"
        f"目前剩餘總名額：{remaining}",
        ephemeral=True
    )

    # 私訊（加入討論串連結）
    try:
        dm_embed = discord.Embed(
            title="🎉 報名成功通知",
            description=f"你已成功報名 **{guild_data['event_name']}**",
            color=discord.Color.green()
        )
        dm_embed.add_field(name="對手幫會", value=guild_data.get("opponent", "未設定"), inline=True)
        dm_embed.add_field(name="約戰時間", value=guild_data.get("event_date", "未設定"), inline=True)
        dm_embed.add_field(name="角色名稱", value=char_name, inline=True)
        dm_embed.add_field(name="職業", value=job, inline=True)
        if thread_link:
            dm_embed.add_field(name="📢 報名討論串", value=f"[點此前往討論串]({thread_link})", inline=False)
        dm_embed.set_footer(text="請準時參加，如有問題請聯繫管理員")
        await interaction.user.send(embed=dm_embed)
    except discord.Forbidden:
        pass

    await asyncio.sleep(8)
    try:
        await interaction.delete_original_response()
    except Exception:
        pass


# ---------- 取消報名 ----------
async def handle_cancel(interaction: discord.Interaction):
    guild = interaction.guild
    if guild is None:
        await interaction.response.send_message("❌ 此功能只能在伺服器內使用。", ephemeral=True)
        return

    guild_data = get_guild_data(guild.id)
    user_id = str(interaction.user.id)

    if is_deadline_passed(guild_data):
        await interaction.response.send_message("⛔ 報名已截止，無法取消報名。", ephemeral=True)
        await asyncio.sleep(5)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass
        return

    if user_id not in guild_data["signups"]:
        await interaction.response.send_message("❌ 你目前沒有報名任何活動。", ephemeral=True)
        await asyncio.sleep(5)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass
        return

    info = guild_data["signups"].pop(user_id)
    update_guild_data(guild.id, guild_data)
    await update_thread_status(guild)

    await interaction.response.send_message(
        f"✅ 已取消報名\n原角色：**{info.get('char_name', '未知')}**　職業：**{info['job']}**",
        ephemeral=True
    )

    try:
        dm_embed = discord.Embed(
            title="❌ 取消報名通知",
            description=f"你已取消 **{guild_data.get('event_name', '活動')}** 的報名",
            color=discord.Color.orange()
        )
        dm_embed.add_field(name="原角色名稱", value=info.get("char_name", "未知"), inline=True)
        dm_embed.add_field(name="原職業", value=info["job"], inline=True)
        await interaction.user.send(embed=dm_embed)
    except discord.Forbidden:
        pass

    await asyncio.sleep(6)
    try:
        await interaction.delete_original_response()
    except Exception:
        pass


# ---------- View ----------
class SignupView(discord.ui.View):
    def __init__(self, disabled: bool = False):
        super().__init__(timeout=None)

        options = [
            discord.SelectOption(label=job, value=job, description=f"報名 {job}")
            for job in JOBS
        ]
        select = discord.ui.Select(
            placeholder="選擇要報名的職業..." if not disabled else "⛔ 報名已截止",
            custom_id="signup_job_select",
            options=options,
            disabled=disabled
        )
        select.callback = self.job_select_callback
        self.add_item(select)

        cancel_btn = discord.ui.Button(
            label="取消報名",
            style=discord.ButtonStyle.danger,
            custom_id="cancel_signup_btn",
            emoji="⚠️",
            disabled=disabled
        )
        cancel_btn.callback = self.cancel_callback
        self.add_item(cancel_btn)

        refresh_btn = discord.ui.Button(
            label="重新整理狀態",
            style=discord.ButtonStyle.secondary,
            custom_id="refresh_status_btn",
            emoji="🔄"
        )
        refresh_btn.callback = self.refresh_callback
        self.add_item(refresh_btn)

    async def job_select_callback(self, interaction: discord.Interaction):
        guild_data = get_guild_data(interaction.guild.id)
        if is_deadline_passed(guild_data):
            await interaction.response.send_message("⛔ 報名已截止。", ephemeral=True)
            await asyncio.sleep(4)
            try:
                await interaction.delete_original_response()
            except Exception:
                pass
            return
        job = interaction.data["values"][0]
        await interaction.response.send_modal(CharacterNameModal(job))

    async def cancel_callback(self, interaction: discord.Interaction):
        await handle_cancel(interaction)

    async def refresh_callback(self, interaction: discord.Interaction):
        await update_thread_status(interaction.guild)
        await interaction.response.send_message("✅ 狀態已重新整理！", ephemeral=True)
        await asyncio.sleep(4)
        try:
            await interaction.delete_original_response()
        except Exception:
            pass


# ---------- 斜線指令 ----------
HOUR_CHOICES = [app_commands.Choice(name=f"{h:02d} 時", value=h) for h in range(24)]
MINUTE_CHOICES = [
    app_commands.Choice(name="00 分（整點）", value=0),
    app_commands.Choice(name="30 分", value=30),
]


@bot.tree.command(name="new_event", description="【管理員】設定聯賽活動並建立討論串")
@app_commands.describe(
    名稱="活動名稱，例如：秋季聯賽",
    對手幫會="對手幫會名稱",
    約戰日期="約戰日期（格式：YYYY-MM-DD，例如 2026-09-20）",
    約戰時="約戰小時",
    約戰分="約戰分鐘（整點或 30 分）",
    截止日期="報名截止日期（格式：YYYY-MM-DD）",
    截止時="報名截止小時",
    截止分="報名截止分鐘（整點或 30 分）"
)
@app_commands.choices(約戰時=HOUR_CHOICES, 約戰分=MINUTE_CHOICES, 截止時=HOUR_CHOICES, 截止分=MINUTE_CHOICES)
@app_commands.default_permissions(administrator=True)
async def set_event(
    interaction: discord.Interaction,
    名稱: str,
    對手幫會: str,
    約戰日期: str,
    約戰時: app_commands.Choice[int],
    約戰分: app_commands.Choice[int],
    截止日期: str,
    截止時: app_commands.Choice[int],
    截止分: app_commands.Choice[int]
):
    await interaction.response.defer(ephemeral=True)

    battle_dt = parse_and_build_datetime(約戰日期, 約戰時.value, 約戰分.value)
    if battle_dt is None:
        await interaction.followup.send(
            "❌ **約戰日期** 格式錯誤！請使用 `YYYY-MM-DD`（例如：2026-09-20）",
            ephemeral=True
        )
        return

    deadline_dt = parse_and_build_datetime(截止日期, 截止時.value, 截止分.value)
    if deadline_dt is None:
        await interaction.followup.send(
            "❌ **報名截止日期** 格式錯誤！請使用 `YYYY-MM-DD`（例如：2026-09-19）",
            ephemeral=True
        )
        return

    if deadline_dt >= battle_dt:
        await interaction.followup.send(
            "❌ 報名截止時間必須**早於**約戰時間！",
            ephemeral=True
        )
        return

    battle_display = format_datetime_display(battle_dt)
    deadline_display = format_datetime_display(deadline_dt)

    guild = interaction.guild
    channel = interaction.channel

    if isinstance(channel, discord.Thread):
        channel = channel.parent
    if not isinstance(channel, discord.TextChannel):
        await interaction.followup.send("❌ 請在文字頻道中使用此指令。", ephemeral=True)
        return

    guild_data = get_guild_data(guild.id)

    guild_data["event_name"] = 名稱
    guild_data["opponent"] = 對手幫會
    guild_data["event_date"] = battle_display
    guild_data["deadline"] = deadline_display
    guild_data["deadline_iso"] = deadline_dt.isoformat()
    guild_data["signups"] = {}
    guild_data["registration_closed"] = False
    guild_data["channel_id"] = channel.id

    try:
        thread = await channel.create_thread(
            name=f"【{名稱}】vs {對手幫會}｜{battle_display}",
            type=discord.ChannelType.public_thread,
            auto_archive_duration=10080
        )
        guild_data["thread_id"] = thread.id

        # 建立後立即鎖定討論串（玩家無法隨意發言，保持乾淨）
        await thread.edit(locked=True)

        embed = format_status_embed(guild_data)
        view = SignupView(disabled=False)

        msg = await thread.send(
            content=(
                f"📢 **{名稱} 報名討論串已建立！**\n"
                f"🏰 對手幫會：{對手幫會}\n"
                f"⏰ 約戰時間：{battle_display}\n"
                f"📅 報名截止：{deadline_display}\n\n"
                f"請使用下方**下拉選單**選擇職業，接著輸入角色名稱完成報名。\n"
                f"• 總人數上限：**{MAX_PLAYERS}** 人\n"
                f"• 每種職業上限：**{MAX_PER_JOB}** 人\n"
                f"• 此討論串已鎖定，僅供查看報名狀態"
            ),
            embed=embed,
            view=view
        )
        guild_data["message_id"] = msg.id
        update_guild_data(guild.id, guild_data)

        await interaction.followup.send(
            f"✅ 活動「**{名稱}**」已成功設定！\n"
            f"對手：{對手幫會}\n"
            f"約戰時間：{battle_display}\n"
            f"報名截止：{deadline_display}\n"
            f"討論串：{thread.mention}（已自動鎖定）",
            ephemeral=True
        )
    except Exception as e:
        await interaction.followup.send(f"❌ 建立討論串失敗：{e}", ephemeral=True)


@bot.tree.command(name="export_excel", description="【管理員】將目前報名名單匯出成 Excel（自動分組）")
@app_commands.default_permissions(administrator=True)
async def export_list(interaction: discord.Interaction):
    await interaction.response.defer(ephemeral=True)

    guild_data = get_guild_data(interaction.guild.id)
    signups = guild_data.get("signups", {})

    if not signups:
        await interaction.followup.send("❌ 目前沒有任何人報名。", ephemeral=True)
        return

    # 智能分組
    groups = auto_group_players(signups)

    wb = Workbook()

    # ===== 工作表1：分組名單（主要） =====
    ws = wb.active
    ws.title = "分組名單"

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    group_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # 標題
    ws.merge_cells('A1:E1')
    title_cell = ws['A1']
    title_cell.value = f"{guild_data.get('event_name', '聯賽')} vs {guild_data.get('opponent', '')}｜{guild_data.get('event_date', '')}"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center")

    # 表頭
    headers = ["組別", "角色名稱", "職業", "Discord名稱", "Discord ID"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    row = 4
    for g_idx, group in enumerate(groups, 1):
        if not group:
            # 空組也顯示
            cell = ws.cell(row=row, column=1, value=f"第 {g_idx} 組")
            cell.fill = group_fill
            cell.border = thin_border
            for c in range(2, 6):
                ws.cell(row=row, column=c).border = thin_border
            row += 1
            continue

        for p_idx, person in enumerate(group):
            values = [
                f"第 {g_idx} 組" if p_idx == 0 else "",
                person["char_name"],
                person["job"],
                person["name"],
                person["uid"]
            ]
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center")
                if p_idx == 0:
                    cell.fill = group_fill
            row += 1

        # 組與組之間空一行
        row += 1

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 22

    # ===== 工作表2：完整名單（未分組） =====
    ws2 = wb.create_sheet("完整名單")
    headers2 = ["序號", "角色名稱", "職業", "Discord名稱", "Discord ID", "報名時間"]
    for col, header in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    sorted_signups = sorted(signups.items(), key=lambda x: (x[1].get("job", ""), x[1].get("char_name", "")))
    for idx, (uid, info) in enumerate(sorted_signups, 1):
        row_data = [
            idx,
            info.get("char_name", "未知"),
            info.get("job", ""),
            info.get("name", ""),
            uid,
            info.get("joined_at", "")[:19].replace("T", " ")
        ]
        for col, value in enumerate(row_data, 1):
            cell = ws2.cell(row=idx + 1, column=col, value=value)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

    for col in range(1, 7):
        ws2.column_dimensions[get_column_letter(col)].width = 18

    # ===== 工作表3：職業統計 =====
    ws3 = wb.create_sheet("職業統計")
    ws3.append(["職業", "目前人數", "上限", "剩餘名額"])
    job_counts = get_job_counts(signups)
    for job in JOBS:
        count = job_counts[job]
        ws3.append([job, count, MAX_PER_JOB, MAX_PER_JOB - count])

    # ===== 工作表4：分組概況 =====
    ws4 = wb.create_sheet("分組概況")
    ws4.append(["組別", "人數", "職業組成"])
    for g_idx, group in enumerate(groups, 1):
        jobs_in_group = [p["job"] for p in group]
        ws4.append([f"第 {g_idx} 組", len(group), "、".join(jobs_in_group) if jobs_in_group else "（空）"])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"{guild_data.get('event_name', '聯賽')}_分組名單_{datetime.now(TZ).strftime('%Y%m%d_%H%M')}.xlsx"
    file = discord.File(fp=buffer, filename=filename)

    total = len(signups)
    await interaction.followup.send(
        content=(
            f"📊 **報名名單已匯出並自動分組**（共 {total} 人）\n"
            f"對手：{guild_data.get('opponent', '')}\n"
            f"已依照「職業不重複 + 人數平均」原則分成 10 組"
        ),
        file=file,
        ephemeral=True
    )


@bot.tree.command(name="event_status", description="查看目前報名狀態")
async def status_cmd(interaction: discord.Interaction):
    guild_data = get_guild_data(interaction.guild.id)
    if not guild_data.get("event_name"):
        await interaction.response.send_message("❌ 目前沒有進行中的活動。", ephemeral=True)
        return
    embed = format_status_embed(guild_data)
    await interaction.response.send_message(embed=embed, ephemeral=True)


@bot.tree.command(name="close_event", description="【管理員】立即關閉報名按鈕")
@app_commands.default_permissions(administrator=True)
async def force_close(interaction: discord.Interaction):
    guild_data = get_guild_data(interaction.guild.id)
    guild_data["registration_closed"] = True
    update_guild_data(interaction.guild.id, guild_data)
    await update_thread_status(interaction.guild)
    await interaction.response.send_message("✅ 已強制關閉報名。", ephemeral=True)


# ---------- 背景任務 ----------
@tasks.loop(minutes=1)
async def check_deadline_task():
    data = load_data()
    for gid, guild_data in data.items():
        if guild_data.get("registration_closed"):
            continue
        if is_deadline_passed(guild_data):
            guild_data["registration_closed"] = True
            update_guild_data(int(gid), guild_data)
            guild = bot.get_guild(int(gid))
            if guild:
                await update_thread_status(guild)
                print(f"⏰ 伺服器 {gid} 報名已自動截止")


@bot.event
async def on_ready():
    print(f"✅ 已登入為 {bot.user} (ID: {bot.user.id})")
    try:
        print(f"⏰ 目前台北時間：{datetime.now(TZ).strftime('%Y-%m-%d %H:%M:%S')}")
    except Exception:
        pass
    bot.add_view(SignupView(disabled=False))
    try:
        synced = await bot.tree.sync()
        print(f"✅ 同步了 {len(synced)} 個斜線指令")
    except Exception as e:
        print(f"❌ 同步斜線指令失敗: {e}")

    if not check_deadline_task.is_running():
        check_deadline_task.start()


if __name__ == "__main__":
    # 引入 keep_alive（用於 Replit 等平台保持 Bot 活著）
    try:
        from keep_alive import keep_alive
        keep_alive()
        print("✅ keep_alive 已啟動")
    except ImportError:
        print("⚠️ 找不到 keep_alive.py，略過 keep_alive（本地執行可忽略）")

    token = os.getenv("DISCORD_TOKEN") or ""
    if token == "":
        raise Exception("Please add your token to the Secrets pane. / 請在 .env 或 Secrets 設定 DISCORD_TOKEN")

    try:
        bot.run(token)
    except discord.HTTPException as e:
        if e.status == 429:
            print("The Discord servers denied the connection for making too many requests")
            print("Get help from https://stackoverflow.com/questions/66724687/in-discord-py-how-to-solve-the-error-for-toomanyrequests")
        else:
            raise e
