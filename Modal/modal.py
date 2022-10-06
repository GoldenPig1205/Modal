from discord.app_commands import Choice
from discord.ui import Button, Select, View
from discord import app_commands, ButtonStyle, SelectOption, TextStyle
from discord.ext import commands
import discord, openpyxl, random, string

bot = commands.Bot(command_prefix="*", intents=discord.Intents.all(), help_command=None)
tree = bot.tree


@bot.event
async def on_ready():
    await tree.sync()
    print("ModalㅣReady")


async def Modal(interaction: discord.Interaction, modal_id):
    def info(key: str):
        openxl = openpyxl.load_workbook("modals.xlsx")
        wb = openxl.active

        for i in range(1, 10001):
            if wb['A' + str(i)].value == modal_id:
                return wb[key + str(i)].value

    def style(key: str):
        if key == "1":
            return TextStyle.short

        elif key == "2":
            return TextStyle.long

    def required(key: str):
        if key == "1":
            return True

        elif key == "2":
            return False

    class Modal(discord.ui.Modal, title=info("B")):
        def desc1():
            if info("C") is not None:
                return discord.ui.TextInput(
                    style=style(info("E")),
                    label=info("C"),
                    placeholder=info("D"),
                    default=info("F"),
                    min_length=int(info("AA").split("/")[0]),
                    max_length=int(info("AA").split("/")[1]),
                    required=required(info("BA"))
                )

            else:
                return None

        desc_1 = desc1()

        def desc2():
            if info("G") is not None:
                return discord.ui.TextInput(
                    style=style(info("I")),
                    label=info("G"),
                    placeholder=info("H"),
                    default=info("J"),
                    min_length=int(info("AB").split("/")[0]),
                    max_length=int(info("AB").split("/")[1]),
                    required=required(info("BB"))
                )

            else:
                return None

        desc_2 = desc2()

        def desc3():
            if info("K") is not None:
                return discord.ui.TextInput(
                    style=style(info("M")),
                    label=info("K"),
                    placeholder=info("L"),
                    default=info("N"),
                    min_length=int(info("AC").split("/")[0]),
                    max_length=int(info("AC").split("/")[1]),
                    required=required(info("BC"))
                )

            else:
                return None

        desc_3 = desc3()

        async def on_submit(self, ctx: discord.Interaction):
            await ctx.response.send_message(f"**📝 모달 작성이 완료되었습니다!**", ephemeral=True)
            embed = discord.Embed(title=f"{bot.user.display_name} / 모달 작성 결과", description=f"**`제목`** : {info('B')}",
                                  color=0x2F3136)
            embed.set_footer(text=f"{interaction.user}님이 작성함", icon_url=interaction.user.display_avatar.url)

            counts = [self.desc_1, self.desc_2, self.desc_3]
            for i in range(0, 3):
                if counts[i] is not None:
                    embed.add_field(name=f"**`부제`** : {counts[i].label}", value=f"**`응답`** : {counts[i].value}",
                                    inline=False)

            await bot.get_channel(int(info("O"))).send(embed=embed)

    await interaction.response.send_modal(Modal())


class 모달(app_commands.Group):
    modal_command = app_commands.Group(name="모달", description="모달을 작동시키기 위한 명령어")

    @app_commands.command(name="만들기", description="모달을 간편하고 쉽게 제작합니다.")
    @app_commands.choices(방식_1=[Choice(name="짧은 응답", value="1"), Choice(name="긴 응답", value="2")],
                          방식_2=[Choice(name="짧은 응답", value="1"), Choice(name="긴 응답", value="2")],
                          방식_3=[Choice(name="짧은 응답", value="1"), Choice(name="긴 응답", value="2")],
                          의무_1=[Choice(name="필수형 응답", value="1"), Choice(name="선택형 응답", value="2")],
                          의무_2=[Choice(name="필수형 응답", value="1"), Choice(name="선택형 응답", value="2")],
                          의무_3=[Choice(name="필수형 응답", value="1"), Choice(name="선택형 응답", value="2")])
    async def create(self, interaction: discord.Interaction, 모달_제목: str, 모달_결과_채널: discord.TextChannel,
                     부제_1: str = "🥚 이스터에그..", 설명_1: str = None, 방식_1: Choice[str] = "1", 미정_1: str = None, 최소_길이_1: int = 0, 최대_길이_1: int = 100, 의무_1: Choice[str] = "1",
                     부제_2: str = None, 설명_2: str = None, 방식_2: Choice[str] = "1", 미정_2: str = None, 최소_길이_2: int = 0, 최대_길이_2: int = 100, 의무_2: Choice[str] = "1",
                     부제_3: str = None, 설명_3: str = None, 방식_3: Choice[str] = "1", 미정_3: str = None, 최소_길이_3: int = 0, 최대_길이_3: int = 100, 의무_3: Choice[str] = "1"):
        await interaction.response.defer(thinking=True, ephemeral=True)

        if not interaction.user.guild_permissions.manage_channels:
            return await interaction.edit_original_message(content=f"**{interaction.guild.name}**에서 **`채널 관리하기`** 권한이 있어야 실행이 가능한 명령어입니다.")
        
        openxl = openpyxl.load_workbook("modals.xlsx")
        wb = openxl.active

        def style(key: Choice[str]):
            try:
                if key.value is None:
                    return key

                else:
                    return key.value
            except AttributeError:
                return key

        def create_id():
            cool_list = []
            for r in range(1, 10):
                cool_list.append(random.choice(string.ascii_letters + string.digits))

            return "".join(cool_list)

        modal_id = create_id()

        for i in range(1, 10001):
            if wb['A' + str(i)].value is None:
                wb['A' + str(i)].value = modal_id
                wb['B' + str(i)].value = 모달_제목
                wb['C' + str(i)].value = 부제_1
                wb['D' + str(i)].value = 설명_1
                wb['E' + str(i)].value = style(방식_1)
                wb['F' + str(i)].value = 미정_1
                wb['G' + str(i)].value = 부제_2
                wb['H' + str(i)].value = 설명_2
                wb['I' + str(i)].value = style(방식_2)
                wb['J' + str(i)].value = 미정_2
                wb['K' + str(i)].value = 부제_3
                wb['L' + str(i)].value = 설명_3
                wb['M' + str(i)].value = style(방식_3)
                wb['N' + str(i)].value = 미정_3
                wb['O' + str(i)].value = str(모달_결과_채널.id)
                wb['Z' + str(i)].value = str(interaction.user.id)
                wb['AA' + str(i)].value = f"{최소_길이_1}/{최대_길이_1}"
                wb['AB' + str(i)].value = f"{최소_길이_2}/{최대_길이_2}"
                wb['AC' + str(i)].value = f"{최소_길이_3}/{최대_길이_3}"
                wb['BA' + str(i)].value = style(의무_1)
                wb['BB' + str(i)].value = style(의무_2)
                wb['BC' + str(i)].value = style(의무_3)
                openxl.save("modals.xlsx")
                break

        view = View()
        view.add_item(Button(label="모달 실행", emoji="📑", style=ButtonStyle.blurple, custom_id=f"모달/{modal_id}"))
        embed = discord.Embed(title=f"{bot.user.display_name} / 모달 만들기", description=f"\✅ **모달이 성공적으로 생성되었습니다.**\n"
                              f"> 모달 아이디 : `{modal_id}`\n> 모달 결과 채널 : <#{모달_결과_채널.id}>", color=0x2F3136)
        embed.set_footer(text=f"모달 아이디는 소중히 보관해 두십시요.", icon_url=interaction.user.display_avatar.url)
        await interaction.edit_original_message(embed=embed, view=view)

    @app_commands.command(name="참여하기", description="모달 아이디를 사용해 빠르게 모달에 참여합니다.")
    @app_commands.describe(모달_아이디="참여할 모달의 아이디를 입력해주세요.")
    async def join(self, interaction: discord.Interaction, 모달_아이디: str):
        openxl = openpyxl.load_workbook("modals.xlsx")
        wb = openxl.active

        for i in range(1, 10001):
            if wb['A' + str(i)].value == 모달_아이디:
                try:
                    return await Modal(interaction, 모달_아이디)
                except Exception as e:
                    await interaction.response.send_message(f"\⚠ **아래 오류가 발생하여 모달을 실행할 수 없습니다.**\n```diff\n- {e}```", ephemeral=True)

            elif wb['A' + str(i)].value is None:
                return await interaction.response.send_message(f"\⚠ **존재하지 않는 `모달 아이디` 입니다.**", ephemeral=True)

    @app_commands.command(name="목록", description="자신이 만든 모달들의 목록을 보여줍니다.")
    async def list_modal(self, interaction: discord.Interaction):
        await interaction.response.defer(thinking=True, ephemeral=True)

        openxl = openpyxl.load_workbook("modals.xlsx")
        wb = openxl.active

        modal_list = []

        for i in range(1, 10001):
            if wb['Z' + str(i)].value == str(interaction.user.id):
                modal_list.append(f"`{wb['A' + str(i)].value}` - **{wb['B' + str(i)].value}**ㅣ<#{wb['O' + str(i)].value}>")

            elif wb['Z' + str(i)].value is None:
                break

        def check_modal():
            if len(modal_list) <= 0:
                return "\⚠ **현재 만들어진 모달이 없습니다.**"

            else:
                return f"\n".join(modal_list)

        view = View()
        view.add_item(Button(label="모달 실행", emoji="📑", style=ButtonStyle.blurple, custom_id=f"process_modal"))
        view.add_item(Button(label="모달 삭제", style=ButtonStyle.red, emoji="🗑", custom_id="delete_modal"))

        embed = discord.Embed(title=f"{bot.user.display_name} / 모달 목록", description=f"`모달 아이디` - **모달 제목**ㅣ모달 결과 채널\n───────────────────\n{check_modal()}", color=0x2F3136)
        embed.set_footer(text=f"{interaction.user.display_name}님의 모달 목록", icon_url=interaction.user.display_avatar.url)
        await interaction.edit_original_message(embed=embed, view=view)


    @app_commands.command(name="문의", description="모달에 관해 문의할 수 있습니다.")
    @app_commands.choices(문의_종류=[Choice(name="⚠ 버그 제보", value="1"), Choice(name="🗳 모달 건의", value="2"),
                                 Choice(name="❗ 악용 사례", value="3"), Choice(name="🎲 그 외 기타", value="4")])
    async def support(self, interaction: discord.Interaction, 문의_종류: Choice[str]):
        class Modal(discord.ui.Modal, title=f"{bot.user.display_name} / 모달 문의"):
            def desc1():
                if 문의_종류.value == "1":
                    return discord.ui.TextInput(
                        style=TextStyle.long,
                        label="모달에 어떤 버그가 발견되었습니까?",
                        placeholder="모달에 관한 버그에 관해 자세하게 설명해 주십시요.",
                        required=True,
                        default=f"버그 명령어 : \n버그에 관한 설명 : ",
                        min_length=10,
                        max_length=500
                    )

                elif 문의_종류.value == "2":
                    return discord.ui.TextInput(
                        style=TextStyle.long,
                        label="모달에 관해 건의하고 싶은 것이 있습니까?",
                        placeholder="모달에 관해 건의하고 싶은 것을 설명해 주십시요.",
                        required=True,
                        default=f"건의하고 싶은 기능 (자세하게) : ",
                        min_length=10,
                        max_length=500
                    )

                elif 문의_종류.value == "3":
                    return discord.ui.TextInput(
                        style=TextStyle.long,
                        label="모달을 어떤식으로 악용할 수 있습니까?",
                        placeholder="모달 악용 사례에 관해 자세하게 설명해 주십시요.",
                        required=True,
                        default=f"악용 사례에 관한 설명 : ",
                        min_length=10,
                        max_length=500
                    )

                elif 문의_종류.value == "4":
                    random_msg = ['다음 모달 업데이트는 언제쯤..?', '모달은 누가 만들었나요?', '디스코드 모달이란 무엇인가요?', '모달 프사는 어디서 구하셨나요?']
                    return discord.ui.TextInput(
                        style=TextStyle.long,
                        label="모달에 관해 하고 싶으신 이야기가 있습니까?",
                        placeholder="모달에 관해 이야기하고 싶은 것을 설명해 주십시요.",
                        required=True,
                        default=f'ex. {random.choice(random_msg)}',
                        min_length=10,
                        max_length=500
                    )

            desc_1 = desc1()

            async def on_submit(self, ctx: discord.Interaction):
                await ctx.response.send_message(f"**\✅ 문의 작성이 완료되었습니다!**\n해당 문의에 관한 답변은 10일내로 {interaction.user.mention}의 DM으로 발송될 것입니다.\n```diff\n- 봇이 DM을 보낼 수 있도록 하십시요.```", ephemeral=True)
                embed = discord.Embed(title=f"{bot.user.display_name} / 모달 문의", description=f"**`작성자`** : {interaction.user.mention}", color=0x2F3136)
                embed.add_field(name=f"**`질문`** : {self.desc_1.label}", value=f"**`응답`** : {self.desc_1.value}", inline=True)
                await bot.get_channel(830246342491111485).send(embed=embed)

        await interaction.response.send_modal(Modal())


    @app_commands.command(name="게시", description="모달 아이디로 해당 모달을 채널에 게시합니다.")
    @app_commands.choices(버튼_방식=[Choice(name="회색", value="회색"), Choice(name="파란색", value="파란색"),
                                Choice(name="빨간색", value="빨간색"), Choice(name="초록색", value="초록색")],
                          임베드_색깔=[Choice(name="빨간색", value=0xFA5858), Choice(name="주황색", value=0xFA8258),
                                  Choice(name="노란색", value=0xF4FA58), Choice(name="초록색", value=0x82FA58),
                                  Choice(name="하늘색", value=0x81DAF5), Choice(name="파란색", value=0x2E64FE),
                                  Choice(name="보라색", value=0x8000FF), Choice(name="분홍색", value=0xFA58F4),
                                  Choice(name="검은색", value=0x000000), Choice(name="살구색", value=0xF6D8CE),
                                  Choice(name="랜덤", value=random.randint(0, 0xFFFFFF))])
    async def post(self, interaction: discord.Interaction, 모달_아이디: str, 게시_채널: discord.TextChannel = None,
                   메세지: str = None, 임베드_설명: str = "여러분, 이 모달에 참여해주세요.", 임베드_색깔: Choice[int] = 0x2F3136, 임베드_사진: str = None, 버튼_메세지: str = "모달 참여",
                   버튼_방식: Choice[str] = ButtonStyle.gray, 버튼_이모지: discord.Emoji = None):
        await interaction.response.defer(thinking=True, ephemeral=True)
        openxl = openpyxl.load_workbook(f"modals.xlsx")
        wb = openxl.active

        for i in range(1, 10001):
            if wb['A' + str(i)].value == str(모달_아이디):
                def post_channel():
                    if 게시_채널 is None:
                        return interaction.channel.id

                    else:
                        return 게시_채널.id

                def button_style

                view = View()
                view.add_item(Button(label=버튼_메세지, style=button_style, emoji=버튼_이모지, custom_id=f"모달/{모달_아이디}"))
                embed = discord.Embed(title=f"{bot.user.display_name} / 모달 게시", description=임베드_설명, color=임베드_색깔)
                embed.set_image(url=임베드_사진)
                embed.set_footer(text=f"{bot.get_user(int(wb['Z' + str(i)].value)).name}님이 만든 모달", icon_url=bot.get_user(int(wb['Z' + str(i)].value)).display_avatar.url)
                await bot.get_channel(post_channel()).send(content=메세지, embed=embed, view=view)
                return await interaction.edit_original_message(content=f"**성공적으로 {게시_채널.mention}에 모달을 게시하였습니다.**")

            elif wb['A' + str(i)].value is None:
                return await interaction.response.send_message(f"\⚠ **해당 모달을 찾을 수 없습니다.**")


tree.add_command(모달())


@bot.event
async def on_interaction(interaction: discord.Interaction):
    if not interaction.type == discord.InteractionType.component:
        return

    openxl = openpyxl.load_workbook(f"modals.xlsx")
    wb = openxl.active

    if interaction.data['custom_id'].startswith("모달"):
        try:
            await Modal(interaction, interaction.data['custom_id'].split("/")[1])
        except Exception as e:
            await interaction.response.send_message(f"\⚠ **아래 오류가 발생하여 모달을 실행할 수 없습니다.**\n```diff\n- {e}```", ephemeral=True)

    elif interaction.data['custom_id'] == "process_modal":
        modal_list = []

        for i in range(1, 10001):
            if wb['Z' + str(i)].value == str(interaction.user.id):
                modal_list.append(SelectOption(
                    label=f"{wb['A' + str(i)].value} - {wb['B' + str(i)].value}ㅣ#{bot.get_channel(int(wb['O' + str(i)].value)).name}",
                    value=wb['A' + str(i)].value))

            elif wb['Z' + str(i)].value is None:
                break

        view = View()
        view.add_item(Select(
            placeholder="실행할 모달을 선택해주세요.",
            options=modal_list,
            custom_id="process_modals"
        ))

        try:
            await interaction.response.send_message(view=view, ephemeral=True)
        except discord.errors.HTTPException:
            await interaction.response.edit_message()

    elif interaction.data['custom_id'] == "process_modals":
        modal_id = interaction.data['values'][0]
        
        try:
            return await Modal(interaction, modal_id)
        except discord.errors.HTTPException:
            return await interaction.response.edit_message(content=f"⚠ **해당 모달을 찾을 수 없었습니다..**", view=None)

    elif interaction.data['custom_id'] == "delete_modal":
        modal_list = []

        for i in range(1, 10001):
            if wb['Z' + str(i)].value == str(interaction.user.id):
                modal_list.append(SelectOption(label=f"{wb['A' + str(i)].value} - {wb['B' + str(i)].value}ㅣ#{bot.get_channel(int(wb['O' + str(i)].value)).name}", value=wb['A' + str(i)].value))

            elif wb['Z' + str(i)].value is None:
                break

        view = View()
        view.add_item(Select(
            placeholder="삭제할 모달을 선택해주세요.",
            options=modal_list,
            custom_id="delete_modals"
        ))

        try:
            await interaction.response.send_message(view=view, ephemeral=True)
        except discord.errors.HTTPException:
            await interaction.response.edit_message()

    elif interaction.data['custom_id'] == "delete_modals":
        modal_id = interaction.data['values'][0]

        for i in range(1, 10001):
            if wb['A' + str(i)].value == str(modal_id):
                wb.delete_rows(i)
                openxl.save("modals.xlsx")
                return await interaction.response.edit_message(content=f"✅ **성공적으로 모달이 삭제되었습니다.**", view=None)

            elif wb['A' + str(i)].value is None:
                return await interaction.response.edit_message(content=f"⚠ **해당 모달을 찾을 수 없었습니다..**", view=None)


bot.run("your token")
