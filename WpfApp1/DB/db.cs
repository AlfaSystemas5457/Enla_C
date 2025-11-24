using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Enla_C.DB
{
    public class ConectDB
    {
        public string server;
        public string baseDatos;
        public string NoEmpresa;
        public string usuario;
        public string password;
        public string RS;
        public string Alm1;
        public string Alm2;
        public bool EditaMP;
        public int tipoBD;

        private static ConectDB instance = null;
        private static IniFile iniFile = null;
        private string connectionString;
        private string iniConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");

        public ConectDB()
        {
            try
            {
                if (!File.Exists(iniConfigPath))
                    File.Create(iniConfigPath).Close();

                iniFile = new IniFile(iniConfigPath);

                LoadSettings();
                BuildConnectionString();
            }
            catch (IOException ioEx)
            {
                throw new Exception("Error de acceso al archivo de configuración.", ioEx);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al inicializar la configuración.", ex);
            }
        }

        public static ConectDB Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ConectDB();
                }
                return instance;
            }
        }

        private void LoadSettings()
        {
            try
            {
                this.server = iniFile.Read("Configuración", "Server");
                this.baseDatos = iniFile.Read("Configuración", "BaseDatos");
                this.NoEmpresa = iniFile.Read("Configuración", "NoEmpresa");
                this.usuario = iniFile.Read("Configuración", "Usuario");
                this.password = iniFile.Read("Configuración", "Password");
                this.RS = iniFile.Read("Configuración", "RS");
                this.Alm1 = iniFile.Read("Configuración", "Alm1");
                this.Alm2 = iniFile.Read("Configuración", "Alm2");
                this.EditaMP = bool.Parse(iniFile.Read("Configuración", "EditaMP"));
                this.tipoBD = int.Parse(iniFile.Read("Configuración", "TipoBD") ?? "0");
            }
            catch (Exception ex)
            {
                throw new Exception("Error al cargar configuración del archivo INI.", ex);
            }
        }

        public void SaveConfig(dbParams dbParams)
        {
            try
            {
                iniFile.Write("Configuración", "Server", dbParams.server);
                iniFile.Write("Configuración", "BaseDatos", dbParams.baseDatos);
                iniFile.Write("Configuración", "NoEmpresa", dbParams.NoEmpresa);
                iniFile.Write("Configuración", "Usuario", dbParams.usuario);
                iniFile.Write("Configuración", "Password", dbParams.password);
                iniFile.Write("Configuración", "RS", dbParams.RS);
                iniFile.Write("Configuración", "Alm1", dbParams.Alm1);
                iniFile.Write("Configuración", "Alm2", dbParams.Alm2);
                iniFile.Write("Configuración", "EditaMP", dbParams.EditaMP.ToString());
                iniFile.Write("Configuración", "TipoBD", dbParams.tipoBD.ToString());
                LoadSettings();
                BuildConnectionString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al guardar la configuración.", ex);
            }
        }

        public void LoadSettingsIni()
        {
            LoadSettings();
            BuildConnectionString();
        }

        private void BuildConnectionString()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"User={usuario};");
                sb.Append($"Password={password};");
                sb.Append($"Database={baseDatos};");
                sb.Append($"DataSource={server};");
                sb.Append("Port=3050;");
                sb.Append("Dialect=3;");
                sb.Append("Charset=UTF8;");
                sb.Append("Pooling=true;MinPoolSize=0;MaxPoolSize=50;");
                connectionString = sb.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al construir el ConnectionString.", ex);
            }
        }

        /// <summary>
        /// Obtiene una conexión abierta. Recuerda cerrarla al finalizar.
        /// </summary>
        public FbConnection GetConnection()
        {
            try
            {
                FbConnection conn = new FbConnection(connectionString);
                conn.Open();
                return conn;
            }
            catch (FbException fbEx)
            {
                throw new Exception("Error al conectar a Firebird. Verifique los datos de conexión.", fbEx);
            }
            catch (Exception ex)
            {
                throw new Exception("Error inesperado al obtener la conexión.", ex);
            }
        }

        /// <summary>
        /// Obtiene respuesta de una prueba de conexión. Recuerda cerrarla al finalizar.
        /// </summary>
        public FbConnection GetConnectionTest()
        {
            try
            {
                return new FbConnection(connectionString);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al crear objeto de conexión.", ex);
            }
        }

        /// <summary>
        /// Ejecuta una consulta SQL que devuelve un DataTable.
        /// </summary>
        public DataTable ExecuteQuery(string sql, params FbParameter[] parameters)
        {
            try
            {
                using (FbConnection conn = GetConnection())
                using (FbCommand cmd = new FbCommand(sql, conn))
                {
                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }

                    using (FbDataAdapter adapter = new FbDataAdapter(cmd))
                    {
                        var dt = new DataTable();
                        adapter.Fill(dt);
                        return dt;
                    }
                }
            }
            catch (FbException fbEx)
            {
                throw new Exception("Error al ejecutar consulta SQL.", fbEx);
            }
        }

        /// <summary>
        /// Ejecuta un comando SQL que no devuelve datos (INSERT, UPDATE, DELETE).
        /// Devuelve número de filas afectadas.
        /// </summary>
        public int ExecuteNonQuery(string sql, params FbParameter[] parameters)
        {
            try
            {
                using (FbConnection conn = GetConnection())
                using (FbCommand cmd = new FbCommand(sql, conn))
                {
                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }
                    return cmd.ExecuteNonQuery();
                }
            }
            catch (FbException fbEx)
            {
                throw new Exception("Error al ejecutar comando SQL.", fbEx);
            }

        }

        /// <summary>
        /// Ejecuta un comando SQL que devuelve un valor escalar (por ejemplo COUNT(*), etc.)
        /// </summary>
        public object ExecuteScalar(string sql, params FbParameter[] parameters)
        {
            try
            {
                using (FbConnection conn = GetConnection())
                using (FbTransaction trans = conn.BeginTransaction())
                using (FbCommand cmd = new FbCommand(sql, conn, trans))
                {
                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }
                    return cmd.ExecuteScalar();
                }
            }
            catch (FbException fbEx)
            {
                throw new Exception("Error al ejecutar comando escalar SQL.", fbEx);
            }

        }
    }
}
