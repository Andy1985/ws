import java.io.IOException;
import java.net.InetSocketAddress;
import java.util.concurrent.Executors;
import org.apache.mina.common.IoAcceptor;
import org.apache.mina.filter.LoggingFilter;
import org.apache.mina.transport.socket.nio.SocketAcceptor;
import org.apache.mina.common.IoServiceConfig;

public class WordServer
{
    private static final int PORT = 9123;

    public static void main(String[] args) throws IOException
    {
        IoAcceptor acceptor = new SocketAcceptor();
        IoServiceConfig cfg = acceptor.getDefaultConfig();

        cfg.getFilterChain().addLast("logger",new LoggingFilter());

        acceptor.bind(new InetSocketAddress(PORT),new WordServerHandler(),cfg);

        System.out.println("word server started.");
    }
}
